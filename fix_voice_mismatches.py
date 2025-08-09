#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
音声IDミスマッチを自動修正するスクリプト
"""

import json
import os
import zipfile
import xml.etree.ElementTree as ET

def read_excel_voice_mapping():
    """Excelファイルから正しい音声IDとテキストのマッピングを読み取り"""
    try:
        xlsx_path = os.path.join("sound", "audio_script.xlsx")
        
        with zipfile.ZipFile(xlsx_path, 'r') as zip_ref:
            # shared strings を読み取り
            shared_strings = []
            try:
                with zip_ref.open('xl/sharedStrings.xml') as strings_file:
                    strings_tree = ET.parse(strings_file)
                    for si in strings_tree.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
                        t = si.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                        if t is not None:
                            shared_strings.append(t.text or '')
                        else:
                            shared_strings.append('')
            except:
                print("共有文字列の読み取りに失敗")
            
            # ワークシートのデータを読み取り
            with zip_ref.open('xl/worksheets/sheet1.xml') as sheet_file:
                sheet_tree = ET.parse(sheet_file)
                
                # テキストから音声IDへのマッピング（逆引き用）
                text_to_voice = {}
                voice_to_text = {}
                
                for row in sheet_tree.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
                    row_data = {}
                    for cell in row.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                        cell_ref = cell.get('r', '')
                        col = ''.join(filter(str.isalpha, cell_ref))
                        
                        v_element = cell.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                        if v_element is not None:
                            cell_type = cell.get('t', '')
                            if cell_type == 's':  # shared string
                                try:
                                    string_index = int(v_element.text)
                                    value = shared_strings[string_index] if string_index < len(shared_strings) else ''
                                except:
                                    value = v_element.text or ''
                            else:
                                value = v_element.text or ''
                            row_data[col] = value
                    
                    # データ行（3行目以降）の処理
                    if 'A' in row_data and 'C' in row_data:
                        voice_id = row_data.get('A', '').strip()
                        text = row_data.get('C', '').strip()
                        if voice_id and text and voice_id.startswith('L'):
                            text_to_voice[text] = voice_id
                            voice_to_text[voice_id] = text
                
                return text_to_voice, voice_to_text
                
    except Exception as e:
        print(f"Excel読み取りエラー: {e}")
        return {}, {}

def fix_voice_mismatches():
    """音声IDのミスマッチを自動修正"""
    print("=== 音声IDミスマッチ自動修正 ===")
    
    # 正しいマッピングを取得
    text_to_voice, voice_to_text = read_excel_voice_mapping()
    print(f"Excel音声マッピング数: {len(text_to_voice)}")
    
    # scenario_voiced.jsonを読み取り
    try:
        with open("scenario_voiced.json", 'r', encoding='utf-8') as f:
            scenario_data = json.load(f)
    except Exception as e:
        print(f"JSONファイル読み取りエラー: {e}")
        return
    
    fixes_made = 0
    total_checked = 0
    
    # 各シナリオデータを修正
    for protagonist_key, protagonist_data in scenario_data.items():
        for block_key, block_data in protagonist_data.items():
            for i, line_data in enumerate(block_data):
                if isinstance(line_data, dict) and 'voice' in line_data and 'text' in line_data:
                    current_voice_id = line_data['voice']
                    scenario_text = line_data['text']
                    
                    # 配列の場合はスキップ（同時再生用）
                    if isinstance(current_voice_id, list):
                        continue
                    
                    total_checked += 1
                    
                    # テキストに対する正しい音声IDを検索
                    if scenario_text in text_to_voice:
                        correct_voice_id = text_to_voice[scenario_text]
                        
                        # 現在の音声IDが間違っている場合
                        if current_voice_id != correct_voice_id:
                            print(f"修正: {protagonist_key}/{block_key}[{i}] {current_voice_id} → {correct_voice_id}")
                            print(f"  テキスト: 「{scenario_text[:50]}{'...' if len(scenario_text) > 50 else ''}」")
                            
                            # 修正実行
                            line_data['voice'] = correct_voice_id
                            fixes_made += 1
                    else:
                        # テキストがExcelにない場合（空のテキストなど）
                        if scenario_text.strip():  # 空でない場合のみ警告
                            print(f"警告: テキストがExcelに見つからない - {protagonist_key}/{block_key}[{i}]")
                            print(f"  テキスト: 「{scenario_text[:50]}{'...' if len(scenario_text) > 50 else ''}」")
    
    print(f"\nチェック対象数: {total_checked}")
    print(f"修正実行数: {fixes_made}")
    
    if fixes_made > 0:
        # 修正されたJSONファイルを保存
        with open("scenario_voiced.json", 'w', encoding='utf-8') as f:
            json.dump(scenario_data, f, ensure_ascii=False, indent=2)
        
        print(f"\n✅ {fixes_made}件の音声IDを修正し、scenario_voiced.jsonを更新しました")
    else:
        print("\n✅ 修正対象がありませんでした")
    
    return fixes_made

if __name__ == "__main__":
    fixes_made = fix_voice_mismatches()