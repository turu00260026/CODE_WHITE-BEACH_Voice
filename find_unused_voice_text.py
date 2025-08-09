#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
未使用音声ファイルに対応するテキストを調査
"""

import json
import os
import re

def read_excel_as_csv():
    """
    以前に変換したCSVデータを再現するため、
    Excelファイルを簡易的に読み込み
    """
    try:
        # zipfileでExcelファイルを読み込み（前回と同じ方法）
        import zipfile
        import xml.etree.ElementTree as ET
        
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
                
                rows_data = []
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
                    
                    if row_data:
                        rows_data.append(row_data)
                
                return rows_data
                
    except Exception as e:
        print(f"Excel読み取りエラー: {e}")
        return []

def get_used_voice_ids():
    """現在使用されている音声IDを取得"""
    try:
        with open("scenario_voiced.json", 'r', encoding='utf-8') as f:
            scenario_data = json.load(f)
        
        used_voice_ids = set()
        
        for protagonist_key, protagonist_data in scenario_data.items():
            for block_key, block_data in protagonist_data.items():
                for line_data in block_data:
                    if isinstance(line_data, dict) and 'voice' in line_data:
                        voice_data = line_data['voice']
                        if isinstance(voice_data, list):
                            for voice_id in voice_data:
                                used_voice_ids.add(voice_id)
                        else:
                            used_voice_ids.add(voice_data)
        
        return used_voice_ids
        
    except Exception as e:
        print(f"エラー: {e}")
        return set()

def get_available_voice_files():
    """利用可能な音声ファイル一覧を取得"""
    audio_dir = "output_audio"
    available_files = set()
    
    if os.path.exists(audio_dir):
        for filename in os.listdir(audio_dir):
            if filename.endswith('.mp3'):
                voice_id = filename.replace('.mp3', '')
                available_files.add(voice_id)
    
    return available_files

def find_unused_voice_text():
    """未使用音声ファイルに対応するテキストを調査"""
    
    print("=== 未使用音声ファイルのテキスト調査 ===")
    
    # データを取得
    excel_data = read_excel_as_csv()
    used_voice_ids = get_used_voice_ids()
    available_files = get_available_voice_files()
    
    # 未使用音声ファイルを特定
    unused_files = available_files - used_voice_ids
    
    print(f"使用中音声ID数: {len(used_voice_ids)}")
    print(f"利用可能ファイル数: {len(available_files)}")
    print(f"未使用ファイル数: {len(unused_files)}")
    
    # Excelデータを辞書に変換
    voice_text_mapping = {}
    
    if excel_data:
        # ヘッダーを確認（2行目が実際のヘッダー）
        if len(excel_data) >= 2:
            headers = []
            header_row = excel_data[1]  # 2行目
            for col in ['A', 'B', 'C', 'D', 'E']:
                if col in header_row:
                    headers.append(header_row[col])
            
            print(f"Excelヘッダー: {headers}")
            
            # データ行を処理（3行目以降）
            for row_data in excel_data[2:]:
                voice_id = row_data.get('A', '').strip()  # id列
                speaker = row_data.get('B', '').strip()   # speaker列
                text = row_data.get('C', '').strip()      # text列
                
                if voice_id and text:
                    voice_text_mapping[voice_id] = {
                        'speaker': speaker,
                        'text': text
                    }
    
    print(f"Excelから取得した音声テキスト数: {len(voice_text_mapping)}")
    
    # 未使用音声ファイルのテキストを調査
    print(f"\n=== 未使用音声ファイルの対応テキスト ===")
    
    unused_with_text = []
    unused_without_text = []
    
    # 未使用ファイルをソート
    unused_sorted = sorted(unused_files, key=lambda x: int(re.findall(r'\d+', x)[0]) if re.findall(r'\d+', x) else 0)
    
    for voice_id in unused_sorted:
        if voice_id in voice_text_mapping:
            voice_info = voice_text_mapping[voice_id]
            unused_with_text.append({
                'voice_id': voice_id,
                'speaker': voice_info['speaker'],
                'text': voice_info['text']
            })
        else:
            unused_without_text.append(voice_id)
    
    print(f"対応テキストあり: {len(unused_with_text)}")
    print(f"対応テキストなし: {len(unused_without_text)}")
    
    if unused_with_text:
        print("\n【未使用だがテキストがある音声ファイル】")
        for i, item in enumerate(unused_with_text, 1):
            print(f"{i:2d}. {item['voice_id']}: {item['speaker']}")
            text_preview = item['text'][:80] + '...' if len(item['text']) > 80 else item['text']
            print(f"    「{text_preview}」")
            print("")
    
    if unused_without_text:
        print("\n【対応テキストが見つからない音声ファイル】")
        for voice_id in unused_without_text:
            print(f"  {voice_id}")
    
    # 詳細ログを保存
    with open('unused_voice_with_text.json', 'w', encoding='utf-8') as f:
        json.dump(unused_with_text, f, ensure_ascii=False, indent=2)
    
    print(f"\n詳細ログを unused_voice_with_text.json に保存しました")
    
    return unused_with_text

if __name__ == "__main__":
    unused_with_text = find_unused_voice_text()