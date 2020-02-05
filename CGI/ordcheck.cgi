#!/usr/bin/ruby
# -*- coding: utf-8 -*-

###メソッドの定義###
# 全角文字をカウント
# 引数：文字列
# 返り値：全角文字の数
def count_multi_byte(string)
  string.each_char.map{|c| c =~ /[ -~｡-ﾟ]/ ? 0 : 1}.reduce(0, &:+) unless (string.ascii_only? || string.match?(/\A[ｧ-ﾝﾞﾟ]+\z/))
end

###メソッドの定義###
# 全角を3,半角を2とした、文字列の長さを導出
# 引数：文字列
# 返り値：文幅
def count_char(string)
  jisu = string.length  #混合文字数
  zenkaku = count_multi_byte(string)  #全角字数のカウント
  if zenkaku == nil
    zenkaku = 0
  end
  hankaku = jisu - zenkaku  #半角字数のカウント
  haba = zenkaku + hankaku * 2 / 3.0
  return haba.ceil  #文幅を返す
end

###ここから###
require("cgi")

c = CGI.new
f = c["copy"]

print("Content-type: text/html; charset=utf-8\n\n")
print("<html>\n  <head>\n")
print("    <title>宣伝文チェック</title>\n")
print("  </head>\n  <body>\n")
print("    <h2>一般企画宣伝文チェック</h2>\n")
mojisu = count_char(f)  #宣伝文の文字数カウント
print("    <p>宣伝文\n      <br>")
print("<input id='copyTarget' size='50' type='text' value='", CGI.escapeHTML(f.to_s),"' readonly>\n")
if mojisu > 30
  print("      <button>字数を減らしてください</button>")
  print("<br>※", CGI.escapeHTML(mojisu.to_s),"字（制限30字）\n")
  print("      <br>    E01: <font color='red'>宣伝文が", CGI.escapeHTML((mojisu-30).to_s), "字超過しています。</font>\n")
else
  print("      <button onclick='copyToClipboard()'>copy</button>")
  print("<br>※", CGI.escapeHTML(mojisu.to_s),"字（制限30字）\n")
  print("      <br>    No Error: 文字数に問題はありませんでした。")
end
print("    </p>\n")
print("    <button type=\"button\" onclick=\"history.back()\">back</button>\n")
print("    <h4>注意：半角は3文字→2字としてカウントします。文字化けチェック機能はありません。</h4>")
print("\n    <h3>この画面では提出したことにはなりません。必ずメールに添付されたリンク先のフォームに提出してください。</h3>")
print("\n  </body>\n
  <script>\n
    function copyToClipboard() {\n
        var copyTarget = document.getElementById(\"copyTarget\");  // コピー対象のテキストを選択する\n
        copyTarget.select();  // 選択しているテキストをクリップボードにコピーする\n
        document.execCommand(\"Copy\");\n
        alert(\"クリップボードにコピーしました！ : \" + copyTarget.value);\n
    }\n
  </script>
  \n</html>")
