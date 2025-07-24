import google.generativeai as genai

# ここに取得したGemini APIキーを貼り付けてください
# APIキーは機密情報のため、実際の運用では環境変数などを用いて安全に管理することを推奨します。
API_KEY = "AIzaSyBoGWeZNwI7emgNasTDu5CZXeTezLNxliA"

# Gemini APIを設定
genai.configure(api_key=API_KEY)

# 使用するGeminiモデルを指定 (サイトの例に合わせ "gemini-2.0-flash" を使用)
# 利用可能なモデルについては、Google AI for Developersのドキュメントを参照してください。
model = genai.GenerativeModel("gemini-2.0-flash")

def get_gemini_response_from_text(prompt_text: str) -> str:
    """
    Gemini APIを使用してテキストを送信し、回答を取得します。

    Args:
        prompt_text: Geminiに送信するテキストプロンプト。

    Returns:
        Geminiからの回答テキスト。
    """
    try:
        # テキストを送信し、回答を生成
        response = model.generate_content(prompt_text)

        # 回答のテキスト部分を取得
        # レスポンスが複数の候補を持つ場合や、テキスト以外の内容を含む場合があるため、
        # `hasattr(response, "text")` でテキストが存在するか確認します。
        if hasattr(response, "text"):
            return response.text
        else:
            return "Geminiからの回答がありませんでした、またはテキスト形式ではありませんでした。"

    except Exception as e:
        return f"Gemini API呼び出し中にエラーが発生しました: {e}"

# 送信するテキストを変数に格納
input_text_variable = "Pythonでデータ分析を行う際によく使われるライブラリを3つ教えてください。"

# Geminiにテキストを送信し、その回答を変数に格納
gemini_response_variable = get_gemini_response_from_text(input_text_variable)

# 結果を表示
print("--- 送信したテキスト ---")
print(input_text_variable)
print("\n--- Geminiからの回答 ---")
print(gemini_response_variable)

# 別のテキストで試す例
another_input_text = "桜の開花時期はいつですか？"
another_gemini_response = get_gemini_response_from_text(another_input_text)

print("\n--- 別の質問 ---")
print(another_input_text)
print("\n--- Geminiからの回答 ---")
print(another_gemini_response)