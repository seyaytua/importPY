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
    
    # カードヘッダー用スタイル（教科書名）
    card_header_style = ParagraphStyle(
        name='CardHeader',
        parent=styles['Heading2'],
        fontName=jp_font,
        fontSize=14,
        leading=18,
        textColor=colors.HexColor('#343a40')
    )
    
    # ラベル用スタイル（ID, PASSWORD）
    label_style = ParagraphStyle(
        name='InfoLabel',
        parent=styles['Normal'],
        fontSize=8,
        leading=10,
        textColor=colors.HexColor('#6c757d'),
        fontName='Helvetica'  # 英語ラベルなのでHelveticaのまま
    )
    
    # 値用スタイル（実際のIDとパスワード）
    value_style = ParagraphStyle(
        name='InfoValue',
        parent=styles['Normal'],
        fontSize=12,
        leading=14,
        textColor=colors.HexColor('#212529'),
        fontName='Courier'  # ID/パスワードは等幅フォントの方が見やすい
    )

    total_files = 0
    for index, row in df.iterrows():
        email = row.iloc[0]
        if not email:
            continue
        
        account_name = email.split('@')[0]
        pdf_filename = sanitize_filename(f"{account_name}_ライセンス情報.pdf")
        pdf_path = os.path.join(output_folder, pdf_filename)

        textbook_data = []
        for i in range(1, len(row), 3):
            if i + 2 < len(row):
                textbook, user_id, password = row.iloc[i], row.iloc[i+1], row.iloc[i+2]
                if pd.notna(user_id) and user_id.strip() and pd.notna(textbook) and textbook.strip():
                    textbook_data.append({"name": textbook, "id": user_id, "pass": password})

        if not textbook_data:
            print(f"スキップ: {email} には有効なデータがありません。")
            continue

        print(f"作成中: {pdf_filename}")
        
        doc = SimpleDocTemplate(
            pdf_path,
            pagesize=A4,
            leftMargin=20*mm,
            rightMargin=20*mm,
            topMargin=20*mm,
            bottomMargin=20*mm
        )
        story = []

        # タイトルと注意書き
        story.append(Paragraph("ライセンス情報", main_title_style))
        story.append(Paragraph("この情報は他の人と共有しないでください", caution_style))
        
        # 各教科書の情報
        for data in textbook_data:
            # ID/パスワードのコンテンツ
            id_content = [
                Paragraph("ID", label_style),
                Spacer(1, 2*mm),
                Paragraph(data['id'], value_style)
            ]
            pass_content = [
                Paragraph("PASSWORD", label_style),
                Spacer(1, 2*mm),
                Paragraph(data['pass'], value_style)
            ]
            
            # ID/パスワードのテーブル
            body_table = Table(
                [[id_content, pass_content]],
                colWidths=[65*mm, 65*mm]
            )
            body_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f8f9fa')),
                ('PADDING', (0,0), (-1,-1), 8),
                ('LEFTPADDING', (0,0), (-1,-1), 10),
                ('RIGHTPADDING', (0,0), (-1,-1), 10),
                ('ROUNDEDCORNERS', (0,0), (-1,-1), 4),
            ]))

            # カード全体のテーブル
            card_table = Table(
                [[Paragraph(data['name'], card_header_style)], [body_table]],
                colWidths=[140*mm]
            )
            card_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('BOTTOMPADDING', (0,0), (0,0), 10),
                ('TOPPADDING', (0,1), (0,1), 5),
                ('LINEBEFORE', (0,0), (-1,-1), 3, colors.HexColor('#17a2b8')),
                ('LEFTPADDING', (0,0), (-1,-1), 15),
            ]))
            
            story.append(card_table)
            story.append(Spacer(1, 8*mm))

        # PDFを生成
        try:
            doc.build(story)
            total_files += 1
        except Exception as e:
            print(f"エラー: {pdf_filename} の作成に失敗しました: {e}")

    print(f"\n処理完了: {total_files}件のPDFファイルを作成しました。")
    print(f"出力先: {output_folder}")


if __name__ == '__main__':
    create_pdf_from_csv()