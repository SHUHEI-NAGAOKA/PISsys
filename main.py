import time
import chromedriver_binary
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException # TimeoutException をインポート


#ログインを実行する関数
def login(user_id,kyoten_id,password,taion):
    #ブラウザを開いて、KEiROWのサイトにアクセス
    global driver 
    driver = webdriver.Chrome()
    driver.get("https://mobile.keirow.com/qr_login.php#")


    #最初のページの読み込みを待って、users_id要素が確認できたら次のステップに進む
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, 'users_id'))
        )
        print("ユーザーIDフィールドの読み込みに成功しました。")
    except TimeoutException:
        print("ユーザーIDフィールドの読み込みに失敗しました。サイトURLを確認してください。")
        driver.quit()
        exit()


    #ログインに必要なユーザーID、拠点ID、パスワード、体温、ログインボタンの要素を取得
    users_id_field = driver.find_element(By.NAME, 'users_id') #ユーザーIDの要素を取得
    kyoten_id_field = driver.find_element(By.NAME, 'kyoten_id') #拠点IDの要素を取得
    password_field = driver.find_element(By.NAME, 'password') #パスワードの要素を取得
    taion_field = driver.find_element(By.NAME, 'taion') #体温入力の要素を取得
    select_taion = Select(taion_field) #体温入力のフィールドをセレクト可能なオブジェクトにする
    element = driver.find_element(By.LINK_TEXT, "日報ログイン") #ログインボタンの要素を取得


    #それぞれの要素に対して引数のログイン情報を入力
    users_id_field.send_keys(user_id)
    kyoten_id_field.send_keys(kyoten_id)
    password_field.send_keys(password)
    select_taion.select_by_value(taion)


    # ログインボタンをクリック
    element.click()


#実行コード
login("10","t22wj","10","36.4")

original_window = driver.current_window_handle
print(f"元のウィンドウハンドル: {original_window}")


try:
    print("新しいウィンドウ（タブ）が開くのを待機中...")
    WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
    print("新しいウィンドウ（タブ）が出現しました。")

    all_window_handles = driver.window_handles
    print(f"取得した全てのウィンドウハンドル: {all_window_handles}")

    new_window_handle = None
    for handle in all_window_handles:
        if handle != original_window:
            new_window_handle = handle
            break

    if new_window_handle:
        driver.switch_to.window(new_window_handle)
        print(f"新しいウィンドウ（ハンドル: {new_window_handle}）に切り替えました。")
    else:
        raise Exception("新しいウィンドウハンドルが見つかりませんでした。")

    print("新しいタブでメインメニューページ読み込みを待機中...")
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'メインメニュー')]"))
    )
    print("メインメニューページへの遷移を確認しました。")

    print("「施術履歴を印刷する」リンクを待機中...")
    # ★★★★ ここが修正されたXPathです ★★★★
    rireki = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, '/search.do') and contains(., '施術履歴を印刷する')]"))
    )
    rireki.click()
    print("「施術履歴を印刷する」リンクをクリックしました。")

except Exception as e:
    print(f"要素の検索またはクリック中にエラーが発生しました: {e}")
    screenshot_path = "error_screenshot_after_login_attempt.png"
    driver.save_screenshot(screenshot_path)
    print(f"エラー時のスクリーンショットを '{screenshot_path}' に保存しました。")
    print("--- 最終的なページのソースコード（新しいタブに切り替え後の場合） ---")
    print(driver.page_source)
    print("--------------------------")
    driver.quit()
    exit()

time.sleep(2)

paitiant_info = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, '/printOnly.do') and contains(., '岡本')]"))
    )
paitiant_info.click()

time.sleep(5)


year_start = driver.find_element(By.NAME, 'optYearStart')
month_start = driver.find_element(By.NAME, 'optMonthStart')
day_start = driver.find_element(By.NAME, 'optDayStart')

year_end = driver.find_element(By.NAME, 'optYearEnd')
month_end = driver.find_element(By.NAME, 'optMonthEnd')
day_end = driver.find_element(By.NAME, 'optDayEnd')

select_year_start = Select(year_start)
select_year_start.select_by_value('2025')
select_month_start = Select(month_start)
select_month_start.select_by_value('6')
select_day_start = Select(day_start)
select_day_start.select_by_value('1')

select_year_end = Select(year_end)
select_year_end.select_by_value('2025')
select_month_end = Select(month_end)
select_month_end.select_by_value('6')
select_day_end = Select(day_end)
select_day_end.select_by_value('31')


output_btn = driver.find_element(By.NAME, 'excel')
output_btn.click()


time.sleep(5)


driver.quit()