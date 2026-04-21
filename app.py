# ========================================
# 工程碳足跡計算系統 - Streamlit 網站版
# ========================================

import streamlit as st
import sys
import os
import tempfile
import shutil
import subprocess
import gdown

# ========================================
# 自動下載資料庫
# ========================================
資料庫_ID = st.secrets.get("GDRIVE_FILE_ID", "")
if not 資料庫_ID:
    st.error("❌ 系統設定錯誤，請聯絡管理員")
    st.stop()
資料庫_路徑 = "資料庫.xlsx"

st.set_page_config(
    page_title="工程碳足跡計算系統",
    page_icon="🌿",
    layout="wide"
)

if not os.path.exists(資料庫_路徑):
    try:
        gdown.download(
            f"https://drive.google.com/uc?id={資料庫_ID}",
            資料庫_路徑,
            quiet=False
        )
    except Exception as e:
        st.error(f"❌ 資料庫載入失敗：{e}")
        st.stop()
st.markdown("""
<style>
  .main-header {
    background: #1B4332;
    color: white;
    padding: 20px 24px;
    border-radius: 10px;
    margin-bottom: 24px;
  }
  .main-header h1 { color: white; font-size: 22px; margin: 0; }
  .main-header p { color: #74C69D; font-size: 13px; margin: 4px 0 0; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
  <h1>🌿 工程碳足跡計算系統</h1>
  <p>山椒魚永續工程股份有限公司 ｜ Formosanus Engineering Sustainable Solutions</p>
</div>
""", unsafe_allow_html=True)

# ========================================
# 初始化 session_state
# ========================================
if '計算完成' not in st.session_state:
    st.session_state['計算完成'] = False
if '結果' not in st.session_state:
    st.session_state['結果'] = {}

# ========================================
# 側邊欄
# ========================================
with st.sidebar:
    if os.path.exists("logo白底.jpg"):
        st.image("logo白底.jpg", width=220)
    st.divider()

    st.subheader("📂 上傳檔案")
    xml_file = st.file_uploader(
        "工程預算書（XML）",
        type=["xml"],
        help="請上傳工程會標準格式的 XML 預算書"
    )

    高信心 = 75
    低信心 = 45

    st.divider()
    st.subheader("📝 上傳者資訊（選填）")
    上傳單位 = st.text_input("公司／單位名稱", placeholder="例：山椒魚永續工程")
    上傳案名 = st.text_input("工程案名稱", placeholder="例：曾文溪護岸工程")
    開始計算 = st.button("▶ 開始計算", type="primary", use_container_width=True)

# ========================================
# 主畫面：未上傳時顯示說明
# ========================================
if not xml_file and not st.session_state['計算完成']:
    col1, col2, col3 = st.columns(3)
    with col1:
        st.info("**步驟一**\n\n📂 在左側上傳工程預算書 XML 檔案")
    with col2:
        st.info("**步驟二**\n\n⚙️ 調整計算設定（可使用預設值）")
    with col3:
        st.info("**步驟三**\n\n▶ 點擊「開始計算」等待結果")

    st.markdown("---")
    st.markdown("### 📋 系統說明")
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("""
**支援的計算方式：**
- ✅ 資料庫模糊比對
- 🤖 AI 語意比對（需 API）
- 👷 人工固定係數（人工類）
- ⬜ 排除項目（管理費等）
        """)
    with col_b:
        st.markdown("""
**產出結果：**
- 📊 互動式儀表板
- 📋 資源清冊 Excel
- 🔍 係數比對結果 Excel
- 📈 碳足跡計算結果 Excel
        """)

# ========================================
# 執行計算
# ========================================
if xml_file and 開始計算:
    st.session_state['計算完成'] = False
    st.session_state['結果'] = {}

    tmp_dir = tempfile.mkdtemp()
    原始目錄 = os.getcwd()

    try:
        # 儲存 XML
        xml_path = os.path.join(tmp_dir, "預算書.xml")
        with open(xml_path, "wb") as f:
            f.write(xml_file.getvalue())

        # 複製資料庫
        if os.path.exists("資料庫.xlsx"):
            shutil.copy("資料庫.xlsx", os.path.join(tmp_dir, "資料庫.xlsx"))
        else:
            st.error("❌ 找不到係數資料庫（資料庫.xlsx），請確認檔案存在")
            st.stop()

        # 複製 Logo
        for logo in ["logo白底.jpg", "logo.png"]:
            if os.path.exists(logo):
                shutil.copy(logo, os.path.join(tmp_dir, logo))

        # 複製 API 設定
        if os.path.exists("設定.txt"):
            shutil.copy("設定.txt", os.path.join(tmp_dir, "設定.txt"))

        # 執行計算
        with st.spinner("⏳ 計算中，請稍候..."):
            result = subprocess.run(
                [sys.executable,
                 os.path.join(原始目錄, "main.py"),
                 "預算書.xml"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                env={**os.environ, "PYTHONIOENCODING": "utf-8"},
                cwd=tmp_dir
            )

        if result.returncode != 0:
            st.error("❌ 計算失敗，請查看詳情")
            with st.expander("錯誤詳情"):
                st.code(result.stderr)
        else:
            st.success("✅ 計算完成！")
            st.session_state['計算完成'] = True

            # 讀取產出檔案存入 session_state
            檔案清單 = os.listdir(tmp_dir)
            for 類型, 後綴 in [
                ("html",  "_儀表板.html"),
                ("清冊",  "_資源清冊.xlsx"),
                ("比對",  "_係數比對.xlsx"),
                ("計算",  "_碳足跡計算.xlsx"),
            ]:
                檔案 = [f for f in 檔案清單 if f.endswith(後綴)]
                if 檔案:
                    with open(os.path.join(tmp_dir, 檔案[0]), "rb") as f:
                        st.session_state['結果'][類型] = (檔案[0], f.read())

            with st.expander("查看計算詳情"):
                st.code(result.stdout)

    except Exception as e:
        st.error(f"發生錯誤：{e}")
        import traceback
        st.code(traceback.format_exc())

    finally:
        os.chdir(原始目錄)
        try:
            shutil.rmtree(tmp_dir)
        except:
            pass

elif xml_file and not 開始計算 and not st.session_state['計算完成']:
    st.success(f"✅ 已上傳：{xml_file.name}")
    st.info("👈 請點左側「開始計算」按鈕")

# ========================================
# 顯示計算結果（session_state 保持結果）
# ========================================
if st.session_state.get('計算完成'):
    結果 = st.session_state.get('結果', {})

    if 'html' in 結果:
        st.markdown("### 📊 碳足跡儀表板")
        st.components.v1.html(
            結果['html'][1].decode('utf-8'),
            height=1200,
            scrolling=True
        )

    st.markdown("### 📥 下載結果檔案")
    col1, col2, col3 = st.columns(3)

    if '清冊' in 結果:
        col1.download_button(
            "📋 資源清冊",
            結果['清冊'][1],
            結果['清冊'][0],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    if '比對' in 結果:
        col2.download_button(
            "🔍 係數比對結果",
            結果['比對'][1],
            結果['比對'][0],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    if '計算' in 結果:
        col3.download_button(
            "📈 碳足跡計算結果",
            結果['計算'][1],
            結果['計算'][0],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )