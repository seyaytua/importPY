import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import re
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from datetime import datetime
import sys

def register_japanese_font():
    """日本語フォントを登録する（macOS/Windows対応）"""
    
    # CIDフォントを使用（日本語対応の組み込みフォント）
    try:
        pdfmetrics.registerFont(UnicodeCIDFont('HeiseiMin-W3'))
        pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
        print("日本語CIDフォントを登録しました。")
        return 'HeiseiKakuGo-W5'
    except Exception as e:
        print(f"CIDフォント登録エラー: {e}")
    
    # TTFフォントの登録を試みる
    font_name_jp = "Japan-Font"
    
    # すでに登録済みかチェック
    if font_name_jp in pdfmetrics.getRegisteredFontNames():
        return font_name_jp

    font_paths = []
    
    # macOSの場合のフォントパス
    if sys.platform == "darwin":
        font_paths = [
            "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
            "/System/Library/Fonts/Hiragino Sans GB.ttc",
            "/Library/Fonts/Arial Unicode.ttf",
            "/System/Library/Fonts/Helvetica.ttc"
        ]
    # Windowsの場合のフォントパス
    elif sys.platform == "win32":
        windir = os.environ.get('WINDIR', 'C:\\Windows')
        font_paths = [
            os.path.join(windir, 'Fonts', 'meiryo.ttc'),
            os.path.join(windir, 'Fonts', 'msgothic.ttc'),
            os.path.join(windir, 'Fonts', 'YuGothM.ttc'),
            os.path.join(windir, 'Fonts', 'YuGothR.ttc')
        ]
    
    # 利用可能なフォントを探す
    for font_path in font_paths:
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont(font_name_jp, font_path))
                print(f"日本語フォント '{font_path}' を登録しました。")
                return font_name_jp
            except Exception as e:
                print(f"フォント '{font_path}' の登録に失敗: {e}")
                continue
    
    print("警告: 日本語TTFフォントが見つかりませんでした。CIDフォントを使用します。")
    return 'HeiseiKakuGo-W5'

def get_file_path():
    """ファイル選択ダイアログを表示し、CSVファイルのパスを取得する"""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="【ステップ1/2】読み込むCSVファイルを選択してください",
        filetypes=[("CSVファイル", "*.csv")]
    )
    return file_path

def get_folder_path():
    """フォルダ選択ダイアログを表示し、出力先フォルダのパスを取得する"""
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(
        title="【ステップ2/2】PDFの出力先フォルダを選択してください"
    )
    return folder_path

def sanitize_filename(filename):
    """ファイル名として使用できない文字をアンダースコアに置換する"""
    return re.sub(r'[\\/*?:"<>|]', '_', filename)

def add_header(canvas, doc, timestamp_str):
    """PDFの各ページにヘッダーを追加する"""
    canvas.saveState()
    
    # 背景ボックスの設定
    box_width = 35*mm
    box_height = 6*mm
    x_position = A4[0] - 20*mm - box_width
    y_position = A4[1] - 12*mm - box_height/2
    
    # 薄いグレーの背景ボックスを描画
    canvas.setFillColor(colors.HexColor('#f8f9fa'))
    canvas.setStrokeColor(colors.HexColor('#dee2e6'))
    canvas.setLineWidth(0.5)
    canvas.roundRect(x_position, y_position, box_width, box_height, 2, fill=1, stroke=1)
    
    # テキストを描画
    canvas.setFillColor(colors.HexColor('#6c757d'))
    canvas.setFont('Helvetica', 8)
    canvas.drawString(
        x_position + 2*mm,
        y_position + 2*mm,
        f"Issue Date: {timestamp_str}"
    )
    
    canvas.restoreState()

def create_pdf_from_csv():
    """メインの処理"""
    jp_font = register_japanese_font()

    csv_path = get_file_path()
    if not csv_path:
        print("CSVファイルが選択されませんでした。処理を中断します。")
        return

    output_folder = get_folder_path()
    if not output_folder:
        print("出力先フォルダが選択されませんでした。処理を中断します。")
        return

    print(f"\nCSVファイル: {csv_path}")
    print(f"出力先フォルダ: {output_folder}\n")

    # タイムスタンプを生成（日付のみ、ドット区切り）
    timestamp_str = datetime.now().strftime("%Y.%m.%d")

    try:
        # エンコーディングを指定して読み込み
        df = pd.read_csv(csv_path, header=None, dtype=str, encoding='utf-8').fillna('')
    except UnicodeDecodeError:
        try:
            # UTF-8で失敗したらShift-JISで試す
            df = pd.read_csv(csv_path, header=None, dtype=str, encoding='shift-jis').fillna('')
        except Exception as e:
            print(f"エラー: CSVファイルの読み込みに失敗しました: {e}")
            return
    except Exception as e:
        print(f"エラー: CSVファイルの読み込みに失敗しました: {e}")
        return

    # 最低2行必要（1行目=タイトル、2行目以降=データ）
    if len(df) < 2:
        print("エラー: CSVファイルに十分なデータがありません。")
        return
    
    # --- スタイルの定義 ---
    styles = getSampleStyleSheet()
    
    # メインタイトル用スタイル
    main_title_style = ParagraphStyle(
        name='MainTitle',
        parent=styles['Title'],
        fontName=jp_font,
        fontSize=18,
        leading=22,
        spaceAfter=6,
        textColor=colors.black,
        alignment=0  # 左揃え
    )
    
    # 注意書き用スタイル
    caution_style = ParagraphStyle(
        name='Caution',
        parent=styles['Normal'],
        fontName=jp_font,
        fontSize=10,
        leading=14,
        spaceAfter=20,
        textColor=colors.HexColor('#d9534f')
    )
    
    # 教科書名用スタイル
    textbook_title_style = ParagraphStyle(
        name='TextbookTitle',
        parent=styles['Heading2'],
        fontName=jp_font,
        fontSize=14,
        leading=18,
        spaceAfter=8,
        textColor=colors.HexColor('#343a40')
    )
    
    # 情報ラベル用スタイル（ID:, PASSWORD:など）
    info_label_style = ParagraphStyle(
        name='InfoLabel',
        parent=styles['Normal'],
        fontSize=10,
        leading=16,
        textColor=colors.black,
        fontName='Helvetica'
    )
    
    # 情報値用スタイル
    info_value_style = ParagraphStyle(
        name='InfoValue',
        parent=styles['Normal'],
        fontSize=11,
        leading=16,
        textColor=colors.black,
        fontName='Courier'
    )

    total_files = 0
    skipped_rows = 0
    
    # 2行目以降を処理（各ユーザーのデータ）
    # 1行目はタイトル行なのでスキップ
    for index in range(1, len(df)):
        row = df.iloc[index]
        email = row.iloc[0]
        
        if not email or not email.strip():
            skipped_rows += 1
            print(f"スキップ: {index+1}行目 - メールアドレスが空です。")
            continue
        
        account_name = email.split('@')[0]
        pdf_filename = sanitize_filename(f"{account_name}_ライセンス情報.pdf")
        pdf_path = os.path.join(output_folder, pdf_filename)

        textbook_data = []
        
        # B列から最後まで4列ずつ処理（教科書名, ID, Password, シリアルコード）
        col_index = 1  # B列から開始（0がA列なので1がB列）
        
        # 行の最後まで処理
        while col_index + 3 < len(row):  # 4列セットが取れる限り続ける
            # 教科書名を取得
            textbook_name = row.iloc[col_index].strip() if row.iloc[col_index] else ""
            
            # ID, Password, シリアルコードを取得
            user_id = row.iloc[col_index + 1].strip() if row.iloc[col_index + 1] else ""
            password = row.iloc[col_index + 2].strip() if row.iloc[col_index + 2] else ""
            serial_code = row.iloc[col_index + 3].strip() if row.iloc[col_index + 3] else ""
            
            # 教科書名とIDの両方がある場合のみ追加
            if textbook_name and user_id:
                textbook_data.append({
                    "name": textbook_name,
                    "id": user_id,
                    "pass": password,
                    "serial": serial_code
                })
            
            col_index += 4  # 次のセットへ（4列進む）

        # 有効なデータがない場合はスキップ
        if not textbook_data:
            skipped_rows += 1
            print(f"スキップ: {email} には有効なデータがありません。")
            continue

        print(f"作成中 ({index}/{len(df)-1}): {pdf_filename} ({len(textbook_data)}件の教科書)")
        
        def add_header_with_timestamp(canvas, doc):
            add_header(canvas, doc, timestamp_str)
        
        doc = SimpleDocTemplate(
            pdf_path,
            pagesize=A4,
            leftMargin=20*mm,
            rightMargin=20*mm,
            topMargin=25*mm,  # ヘッダー分のスペースを確保
            bottomMargin=20*mm
        )
        story = []

        # タイトルと注意書き
        story.append(Paragraph("ライセンス情報", main_title_style))
        story.append(Paragraph("この情報は他の人と共有しないでください", caution_style))
        
        # 各教科書の情報
        for data in textbook_data:
            # 教科書名
            story.append(Paragraph(f"教科書：{data['name']}", textbook_title_style))
            
            # 情報テーブルのデータを作成
            info_rows = []
            
            # IDは必須（ここまで来ている時点で必ず存在）
            info_rows.append([
                Paragraph("ID:", info_label_style),
                Paragraph(data['id'], info_value_style)
            ])
            
            # パスワードがある場合
            if data['pass']:
                info_rows.append([
                    Paragraph("PASSWORD:", info_label_style),
                    Paragraph(data['pass'], info_value_style)
                ])
            
            # シリアルコードがある場合
            if data['serial']:
                info_rows.append([
                    Paragraph("SERIAL CODE:", info_label_style),
                    Paragraph(data['serial'], info_value_style)
                ])
            
            # 情報のテーブル
            info_table = Table(
                info_rows,
                colWidths=[35*mm, 115*mm]
            )
            info_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f8f9fa')),
                ('PADDING', (0,0), (-1,-1), 8),
                ('LEFTPADDING', (0,0), (0,-1), 10),
                ('ALIGN', (0,0), (0,-1), 'LEFT'),
                ('ALIGN', (1,0), (1,-1), 'LEFT'),
            ]))
            
            story.append(info_table)
            story.append(Spacer(1, 10*mm))

        # PDFを生成（ヘッダー付き）
        try:
            doc.build(story, onFirstPage=add_header_with_timestamp, onLaterPages=add_header_with_timestamp)
            total_files += 1
        except Exception as e:
            print(f"エラー: {pdf_filename} の作成に失敗しました: {e}")

    print(f"\n{'='*60}")
    print(f"処理完了:")
    print(f"  作成成功: {total_files}件のPDFファイル")
    print(f"  スキップ: {skipped_rows}行")
    print(f"  出力先: {output_folder}")
    print(f"{'='*60}")


if __name__ == '__main__':
    create_pdf_from_csv()