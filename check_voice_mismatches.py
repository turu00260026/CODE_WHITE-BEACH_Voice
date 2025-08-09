#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
音声IDの割り当てミスマッチを確認するスクリプト
"""

import json
import os
import zipfile
import xml.etree.ElementTree as ET

def read_excel_voice_mapping():
    """Excelファイルから音声IDとテキストのマッピングを読み取り"""
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
                
                voice_mapping = {}
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
                            voice_mapping[voice_id] = text
                
                return voice_mapping
                
    except Exception as e:
        print(f"Excel読み取りエラー: {e}")
        return {}

def check_voice_mismatches():
    """音声IDとテキストのミスマッチを確認"""
    print("=== 音声ID割り当てミスマッチチェック ===")
    
    # Excelから正しいマッピングを読み取り
    excel_mapping = read_excel_voice_mapping()
    print(f"Excel音声マッピング数: {len(excel_mapping)}")
    
    # scenario_voiced.jsonを読み取り
    try:
        with open("scenario_voiced.json", 'r', encoding='utf-8') as f:
            scenario_data = json.load(f)
    except Exception as e:
        print(f"JSONファイル読み取りエラー: {e}")
        return
    
    mismatches = []
    total_checked = 0
    
    # 各シナリオデータをチェック
    for protagonist_key, protagonist_data in scenario_data.items():
        for block_key, block_data in protagonist_data.items():
            for i, line_data in enumerate(block_data):
                if isinstance(line_data, dict) and 'voice' in line_data and 'text' in line_data:
                    voice_id = line_data['voice']
                    scenario_text = line_data['text']
                    
                    # 配列の場合はスキップ（同時再生用）
                    if isinstance(voice_id, list):
                        continue
                    
                    if voice_id in excel_mapping:
                        excel_text = excel_mapping[voice_id]
                        total_checked += 1
                        
                        # テキストが一致しない場合
                        if scenario_text != excel_text:
                            mismatches.append({
                                'voice_id': voice_id,
                                'location': f"{protagonist_key}/{block_key}[{i}]",
                                'scenario_text': scenario_text,
                                'excel_text': excel_text
                            })
    
    print(f"チェック対象数: {total_checked}")
    print(f"ミスマッチ数: {len(mismatches)}")
    
    if mismatches:
        print("\n【ミスマッチが検出されました】")
        for i, mismatch in enumerate(mismatches, 1):
            print(f"\n{i}. 音声ID: {mismatch['voice_id']}")
            print(f"   場所: {mismatch['location']}")
            print(f"   シナリオ: 「{mismatch['scenario_text'][:100]}{'...' if len(mismatch['scenario_text']) > 100 else ''}」")
            print(f"   Excel:   「{mismatch['excel_text'][:100]}{'...' if len(mismatch['excel_text']) > 100 else ''}」")
    else:
        print("\n✅ ミスマッチは検出されませんでした")
    
    # 詳細ログを保存
    if mismatches:
        with open('voice_mismatches.json', 'w', encoding='utf-8') as f:
            json.dump(mismatches, f, ensure_ascii=False, indent=2)
        print(f"\n詳細ログを voice_mismatches.json に保存しました")
    
    return mismatches

if __name__ == "__main__":
    mismatches = check_voice_mismatches()