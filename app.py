import streamlit as st
import pandas as pd
import datetime
from shift_optimizer import solve
import os

st.set_page_config(page_title="シフト最適化システム", page_icon="📅", layout="wide")

st.title("📅 シフト最適化システム")
st.markdown("---")

# 現在のフォルダパスを取得
current_dir = os.getcwd()

# サイドバーで入力ファイルの確認
with st.sidebar:
    st.header("ℹ️ システム情報")
    st.info(f"""
    このアプリは、スタッフのシフトを自動最適化します。
    
    **作業フォルダ:**
    `{current_dir}`
    
    **必要なファイル:**
    - Shift_Input.xlsx
    - constraints.yaml
    """)
    
    # ファイルの存在確認
    input_file_exists = os.path.exists('Shift_Input.xlsx')
    constraints_file_exists = os.path.exists('constraints.yaml')
    
    st.subheader("📁 ファイルチェック")
    if input_file_exists:
        st.success("✅ Shift_Input.xlsx")
        # ファイル情報を表示
        file_stat = os.stat('Shift_Input.xlsx')
        file_time = datetime.datetime.fromtimestamp(file_stat.st_mtime)
        st.caption(f"更新日時: {file_time.strftime('%Y/%m/%d %H:%M')}")
    else:
        st.error("❌ Shift_Input.xlsx が見つかりません")
    
    if constraints_file_exists:
        st.success("✅ constraints.yaml")
        # ファイル情報を表示
        file_stat = os.stat('constraints.yaml')
        file_time = datetime.datetime.fromtimestamp(file_stat.st_mtime)
        st.caption(f"更新日時: {file_time.strftime('%Y/%m/%d %H:%M')}")
    else:
        st.error("❌ constraints.yaml が見つかりません")

# タブを作成
tab1, tab2, tab3 = st.tabs(["🚀 シフト最適化", "📋 設定ファイル確認", "⚙️ 制約条件設定"])

# タブ1: シフト最適化
with tab1:
    st.subheader("🚀 シフト最適化を実行")
    st.write("下のボタンをクリックして、シフトの最適化を開始します。")
    
    if st.button("▶️ シフトを最適化する", type="primary", use_container_width=True):
        if not input_file_exists or not constraints_file_exists:
            st.error("⚠️ 必要なファイルが見つかりません。ファイルを確認してください。")
        else:
            with st.spinner("🔄 シフトを最適化中... しばらくお待ちください..."):
                try:
                    # シフト最適化を実行
                    solve()
                    st.success("✅ シフトの最適化が完了しました！")
                    st.balloons()
                    
                    # 結果ファイルの確認
                    if os.path.exists('shift_schedule_final.xlsx'):
                        st.info("📄 結果ファイル: `shift_schedule_final.xlsx` が作成されました。")
                        
                        # 結果をプレビュー
                        try:
                            df_result = pd.read_excel('shift_schedule_final.xlsx', sheet_name='シフト表')
                            st.subheader("📊 シフト表プレビュー")
                            # データ型を文字列に変換してArrowエラーを回避
                            df_result = df_result.astype(str)
                            st.dataframe(df_result, use_container_width=True, height=400)
                            
                            # ダウンロードボタン
                            with open('shift_schedule_final.xlsx', 'rb') as f:
                                st.download_button(
                                    label="⬇️ シフト表をダウンロード",
                                    data=f,
                                    file_name=f"shift_schedule_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                        except Exception as e:
                            st.warning(f"プレビューの表示に失敗しました: {e}")
                    
                except Exception as e:
                    st.error(f"❌ エラーが発生しました: {str(e)}")
                    st.exception(e)
    
    # 使い方
    with st.expander("📝 使い方"):
        st.markdown("""
        1. **Shift_Input.xlsx** を準備
           - 「設定」シート: 年月と人数設定
           - 「休日回数」シート: スタッフ情報
           - 「希望休」シート: 休日希望
        
        2. **constraints.yaml** を設定
           - シフト期間
           - 祝日リスト
           - 制約条件
        
        3. 「シフトを最適化する」ボタンをクリック
        
        4. 結果をダウンロード
        """)

# タブ2: 設定ファイル確認
with tab2:
    st.subheader("📋 設定ファイルの内容確認")
    
    if os.path.exists('Shift_Input.xlsx'):
        st.markdown("### 📊 Shift_Input.xlsx")
        
        # 各シートを読み込んで表示
        excel_file = pd.ExcelFile('Shift_Input.xlsx')
        sheet_names = excel_file.sheet_names
        
        selected_sheet = st.selectbox("表示するシートを選択", sheet_names)
        
        if selected_sheet == '設定':
            df = pd.read_excel('Shift_Input.xlsx', sheet_name=selected_sheet, header=None)
            # データ型を文字列に変換してArrowエラーを回避
            df = df.astype(str)
            st.dataframe(df, use_container_width=True)
            st.info("""
            **設定シートの構造:**
            - A1: 年, B1: 値
            - A2: 月, B2: 値
            - A4: 平日最低人数, B4: 値
            - A5: 平日最大人数, B5: 値
            - A6: 休日最低人数, B6: 値
            - A7: 休日最大人数, B7: 値
            - D列: 祝日リスト（日付形式）
            """)
        elif selected_sheet == '休日回数':
            df = pd.read_excel('Shift_Input.xlsx', sheet_name=selected_sheet)
            # データ型を文字列に変換
            df = df.astype(str)
            st.dataframe(df, use_container_width=True)
            st.info("スタッフの基本情報、月間休日数、勤務時間が設定されています。")
        elif selected_sheet == '希望休':
            df = pd.read_excel('Shift_Input.xlsx', sheet_name=selected_sheet, header=0)
            # データ型を文字列に変換
            df = df.astype(str)
            st.dataframe(df, use_container_width=True)
            st.info("各スタッフの希望休が「希」「有」「休」で記載されています。")
        else:
            df = pd.read_excel('Shift_Input.xlsx', sheet_name=selected_sheet)
            # データ型を文字列に変換
            df = df.astype(str)
            st.dataframe(df, use_container_width=True)
    else:
        st.warning("⚠️ Shift_Input.xlsx が見つかりません")
    
    st.markdown("---")
    
    if os.path.exists('constraints.yaml'):
        st.markdown("### 📄 constraints.yaml")
        with open('constraints.yaml', 'r', encoding='utf-8') as f:
            yaml_content = f.read()
        st.code(yaml_content, language='yaml')
    else:
        st.warning("⚠️ constraints.yaml が見つかりません")

# タブ3: 制約条件設定
with tab3:
    st.subheader("⚙️ 制約条件の編集")
    
    if os.path.exists('constraints.yaml'):
        import yaml
        
        with open('constraints.yaml', 'r', encoding='utf-8') as f:
            constraints = yaml.safe_load(f)
        
        st.markdown("### 📝 YAML設定の編集")
        
        # YAMLをテキストエリアで編集可能にする
        edited_yaml = st.text_area(
            "constraints.yaml の内容を編集",
            value=open('constraints.yaml', 'r', encoding='utf-8').read(),
            height=400
        )
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("💾 保存", type="primary"):
                try:
                    # YAML形式が正しいか検証
                    yaml.safe_load(edited_yaml)
                    with open('constraints.yaml', 'w', encoding='utf-8') as f:
                        f.write(edited_yaml)
                    st.success("✅ constraints.yaml を保存しました")
                    st.rerun()
                except yaml.YAMLError as e:
                    st.error(f"❌ YAML形式エラー: {e}")
        
        with col2:
            if st.button("🔄 元に戻す"):
                st.rerun()
        
        st.markdown("---")
        st.info("""
        **設定項目の説明:**
        - `period.start_day`: シフト開始日
        - `period.end_day`: シフト終了日
        - `period.holidays`: 祝日リスト（YYYY-MM-DD形式）※ 現在は設定シートD列から読み込みます
        - `rules.max_consecutive_work_days`: 最大連続勤務日数
        """)
    else:
        st.warning("⚠️ constraints.yaml が見つかりません")

# フッター
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray; font-size: 12px;'>
    シフト最適化システム v1.0 | OR-Tools & Streamlit
    </div>
    """,
    unsafe_allow_html=True
)

# 既存の結果ファイルを表示
if os.path.exists('shift_schedule_final.xlsx') and not st.session_state.get('just_created'):
    st.markdown("---")
    st.subheader("📂 以前の結果")
    try:
        df_existing = pd.read_excel('shift_schedule_final.xlsx', sheet_name='シフト表')
        with st.expander("以前のシフト表を表示"):
            # データ型を文字列に変換
            df_existing = df_existing.astype(str)
            st.dataframe(df_existing, use_container_width=True, height=300)
            with open('shift_schedule_final.xlsx', 'rb') as f:
                st.download_button(
                    label="⬇️ この結果をダウンロード",
                    data=f,
                    file_name="shift_schedule_final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except:
        pass
