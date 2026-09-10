import streamlit as st
import pandas as pd
import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io
from datetime import datetime
import os

# ページ設定
st.set_page_config(
    page_title="アパレルプリント受発注管理システム",
    page_icon="👕",
    layout="wide"
)

# セッション状態の初期化（簡易データストア）
if "inventory" not in st.session_state:
    st.session_state.inventory = pd.DataFrame([
        {"商品ID": "B001", "ボディ名": "スタンダードTシャツ", "カラー": "ホワイト", "サイズ": "L", "単価": 1200, "在庫数": 50},
        {"商品ID": "B002", "ボディ名": "ヘビーウェイトパーカー", "カラー": "ブラック", "サイズ": "XL", "単価": 3500, "在庫数": 20},
    ])

if "orders" not in st.session_state:
    st.session_state.orders = []

# --- サイドバーナビゲーション ---
st.sidebar.title("👕 受発注管理メニュー")
menu = st.sidebar.selectbox(
    "機能を選択してください",
    ["📊 ダッシュボード・在庫管理", "📝 見積り・注文作成", "📄 見積書・請求書PDF出力", "💰 利益計算・売上履歴"]
)

# -------------------------------------------------------------------------
# 1. ダッシュボード・在庫管理
# -------------------------------------------------------------------------
if menu == "📊 ダッシュボード・在庫管理":
    st.title("📊 在庫管理 & ダッシュボード")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("登録ボディ種類", len(st.session_state.inventory))
    with col2:
        total_stock = st.session_state.inventory["在庫数"].sum()
        st.metric("総在庫数", f"{total_stock} 着")
    with col3:
        total_orders = len(st.session_state.orders)
        st.metric("総受注件数", f"{total_orders} 件")
    
    st.markdown("---")
    st.subheader("📦 現在の在庫一覧・編集")
    
    edited_df = st.data_editor(
        st.session_state.inventory, 
        num_rows="dynamic",
        key="inventory_editor"
    )
    st.session_state.inventory = edited_df

# -------------------------------------------------------------------------
# 2. 見積り・注文作成
# -------------------------------------------------------------------------
elif menu == "📝 見積り・注文作成":
    st.title("📝 新規見積り・注文作成")
    
    with st.form("order_form"):
        col1, col2 = st.columns(2)
        with col1:
            client_name = st.text_input("顧客名 / チーム名", "〇〇バスケットボールチーム 様")
            order_date = st.date_input("注文日", datetime.today())
        with col2:
            tax_rate = st.selectbox("消費税率", [0.10, 0.08], format_func=lambda x: f"{int(x*100)}%")
            shipping_fee = st.number_input("送料・手数料 (円)", value=1000, step=100)
            
        st.markdown("---")
        st.subheader("👕 注文アイテム選択")
        
        inventory_df = st.session_state.inventory
        body_options = inventory_df["ボディ名"] + " (" + inventory_df["カラー"] + " / " + inventory_df["サイズ"] + ")"
        
        selected_item_str = st.selectbox("ボディ選択", body_options)
        selected_row = inventory_df[body_options == selected_item_str].iloc[0]
        
        col_q1, col_q2, col_q3 = st.columns(3)
        with col_q1:
            quantity = st.number_input("枚数", min_value=1, value=10)
        with col_q2:
            unit_price = st.number_input("販売単価 (円/枚)", value=int(selected_row["単価"] * 1.5))
        with col_q3:
            print_cost = st.number_input("プリント加工費 (円/枚)", value=500)
            
        submitted = st.form_submit_button("🛒 注文リストに追加")
        
        if submitted:
            subtotal = (unit_price + print_cost) * quantity
            new_order = {
                "注文日時": order_date.strftime("%Y-%m-%d"),
                "顧客名": client_name,
                "商品名": selected_row["ボディ名"],
                "仕様": f"{selected_row['カラー']} / {selected_row['サイズ']}",
                "枚数": quantity,
                "単価": unit_price + print_cost,
                "小計": subtotal,
                "送料": shipping_fee,
                "税率": tax_rate
            }
            st.session_state.orders.append(new_order)
            st.success("注文リストに追加しました！")

    if st.session_state.orders:
        st.markdown("---")
        st.subheader("📋 現在セッションの注文一覧")
        orders_df = pd.DataFrame(st.session_state.orders)
        st.dataframe(orders_df, use_container_width=True)
        
        if st.button("🗑️ 注文履歴をクリア"):
            st.session_state.orders = []
            st.rerun()

# -------------------------------------------------------------------------
# 3. 見積書・請求書PDF出力
# -------------------------------------------------------------------------
elif menu == "📄 見積書・請求書PDF出力":
    st.title("📄 PDF書類出力 (見積書 / 請求書)")
    
    if not st.session_state.orders:
        st.warning("出力する注文データがありません。「見積り・注文作成」タブから注文を追加してください。")
    else:
        client_options = list(set([o["顧客名"] for o in st.session_state.orders]))
        selected_client = st.selectbox("出力対象の顧客を選択", client_options)
        
        client_orders = [o for o in st.session_state.orders if o["顧客名"] == selected_client]
        
        st.write(f"**対象顧客:** {selected_client}")
        client_df = pd.DataFrame(client_orders)
        st.dataframe(client_df, use_container_width=True)
        
        if st.button("📥 御見積書 PDF生成"):
            try:
                # PDFバッファの作成
                pdf_buffer = io.BytesIO()
                doc = SimpleDocTemplate(
                    pdf_buffer, 
                    pagesize=A4,
                    rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40
                )
                elements = []
                
                # --- 【対策1適用部分】クラウド環境でも安全にフォールバックする仕組み ---
                font_path = "C:/Windows/Fonts/meiryo.ttc"
                base_font = 'Helvetica' # デフォルト（クラウド用安全フォント）
                
                if os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(TTFont('Meiryo', font_path))
                        base_font = 'Meiryo'
                    except Exception:
                        pass
                # -----------------------------------------------------------------
                
                styles = getSampleStyleSheet()
                title_style = ParagraphStyle(
                    'TitleStyle',
                    parent=styles['Normal'],
                    fontName=base_font,
                    fontSize=18,
                    leading=22,
                    alignment=1 # 中央揃え
                )
                normal_style = ParagraphStyle(
                    'NormalStyle',
                    parent=styles['Normal'],
                    fontName=base_font,
                    fontSize=10,
                    leading=14
                )
                
                # タイトル
                elements.append(Paragraph("御 見 積 書", title_style))
                elements.append(Spacer(1, 20))
                
                # 宛先・日付
                file_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                header_data = [
                    [Paragraph(f"<b>宛先:</b> {selected_client}", normal_style), Paragraph(f"<b>発行日:</b> {datetime.today().strftime('%Y年%m月%d日')}", normal_style)],
                    [Paragraph("<b>発行元:</b> アプリプリント工房", normal_style), Paragraph(f"<b>管理番号:</b> Q-{file_timestamp}", normal_style)]
                ]
                t_header = Table(header_data, colWidths=[270, 260])
                t_header.setStyle(TableStyle([
                    ('VALIGN', (0,0), (-1,-1), 'TOP'),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 10),
                ]))
                elements.append(t_header)
                elements.append(Spacer(1, 15))
                
                # 明細テーブルデータの作成
                table_data = [["商品名・仕様", "数量", "単価 (円)", "小計 (円)"]]
                total_amount = 0
                shipping = client_orders[0].get("送料", 1000) if client_orders else 1000
                tax_rate = client_orders[0].get("税率", 0.10) if client_orders else 0.10
                
                for item in client_orders:
                    table_data.append([
                        Paragraph(f"{item['商品名']} ({item['仕様']})", normal_style),
                        str(item['枚数']),
                        f"{item['単価']:,}",
                        f"{item['小計']:,}"
                    ])
                    total_amount += item['小計']
                
                # 送料・税金の行を追加
                table_data.append([Paragraph("送料・梱包費", normal_style), "1", f"{shipping:,}", f"{shipping:,}"])
                sub_total_with_shipping = total_amount + shipping
                consumption_tax = int(sub_total_with_shipping * tax_rate)
                grand_total = sub_total_with_shipping + consumption_tax
                
                table_data.append(["", "", "小計", f"{sub_total_with_shipping:,}"])
                table_data.append(["", "", f"消費税 ({int(tax_rate*100)}%)", f"{consumption_tax:,}"])
                table_data.append(["", "", "合計金額", f"{grand_total:,}"])
                
                t_sum = Table(table_data, colWidths=[230, 60, 120, 120])
                t_sum.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2E6EA5')),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                    ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('GRID', (0,0), (-1,-5), 0.5, colors.grey),
                    ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#2E6EA5')),
                    ('TOPPADDING', (0,0), (-1,-1), 6),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                ]))
                elements.append(t_sum)
                
                doc.build(elements)
                
                st.download_button(
                    label="📥 御見積書PDFをダウンロード",
                    data=pdf_buffer.getvalue(),
                    file_name=f"御見積書_{selected_client}_{file_timestamp}.pdf",
                    mime="application/pdf"
                )
                st.success("PDFの生成に成功しました！")
                
            except Exception as e:
                st.error(f"PDF生成エラーが発生しました: {e}")

# -------------------------------------------------------------------------
# 4. 利益計算・売上履歴
# -------------------------------------------------------------------------
elif menu == "💰 利益計算・売上履歴":
    st.title("💰 利益計算 & 売上履歴")
    
    if not st.session_state.orders:
        st.info("まだ受注データがありません。")
    else:
        df_history = pd.DataFrame(st.session_state.orders)
        
        total_sales = df_history["小計"].sum()
        st.metric("総売上金額", f"¥ {total_sales:,}")
        
        st.subheader("📊 受注履歴データ")
        st.dataframe(df_history, use_container_width=True)
