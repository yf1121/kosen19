#encoding: Windows-31J
#掲載情報のデータをチェックし、統合するプログラム
###メソッドの定義###
# 全角文字をカウント
# 引数：文字列
# 返り値：全角文字の数
def count_multi_byte(string)
  string.each_char.map{|c| c.bytesize == 1 ? 0 : 1}.reduce(0, &:+) unless string.ascii_only?
end

###メソッドの定義###
# 全角を3,半角を2とした、文字列の長さを導出
# 引数：文字列
# 返り値：文幅
def count_char(string)
  if string == nil
    haba = 0
  else
    jisu = string.length  #混合文字数
    zenkaku = count_multi_byte(string)  #全角字数のカウント
    if zenkaku == nil
      zenkaku = 0
    end
    hankaku = jisu - zenkaku  #半角字数のカウント
    haba = zenkaku + hankaku * 2 / 3.0
  end
  return haba.ceil  #文幅を返す
end

###メソッドの定義###
# 文字化けの可能性があるか判定
# 引数：文字列, id, 学術紹介文か
# 返り値：エラーコード
def emoji_opt(text, id, lab)
  out = ""
  if text.index("・") == 0 || text.rindex("・") == text.length
    disp = "*90%"
    out = "E04a"
  elsif text.include?("・")
    r = text.split("・")
    k = 1
    while k < r.size
      str = r[k].slice(0,2)
      if str.scan(/[\p{Han}\p{Katakana}]+/).uniq.length == 0 || r[k].length == 0
        disp = "*70%"
        out = "E04b"
        break
      end
      k += 1
    end
    if text.index("・") == text.rindex("・") && (text.index("・") - text.length / 2) > -(text.length / 6)
      disp = "*40%"
      out = "E04c"
    elsif k == r.size
      disp = "*10%"
      out = "E04d"
    end
  end
  if out.length > 0
    if lab == true
      print("    ", out, "-s: 文字化け疑惑(学術紹介文)", disp,"(id=", id, ")\n")
      out = ";" + out + "-s"
    else
      print("    ", out, ": 文字化け疑惑(宣伝文)", disp,"(id=", id, ")\n")
      out = ";" + out
    end
  end
  return out
end

###メソッドの定義###
# 全角を3,半角を2とした、文字列の長さを導出
# 引数：配列
# 返り値：6つのタグアイコン
def tag_judge(array)
  tag6 = Array.new
  tag19 = ["ミニ販売コーナー", "ミニ体験コーナー", "ミニ展示コーナー",
      "子供も楽しめる", "休憩できる", "おひとりさま歓迎", "撮影禁止",
      "フラッシュ撮影禁止", "ビデオ撮影・録音禁止", "展示作品の接触禁止",
      "無許可SNS投稿禁止", "飲食禁止", "テイクアウト可", "途中入退場可",
      "時間入れ替え制", "整理券配布あり", "先着順", "有料", "カンパ制"]
  tagname = ["11_ミニ販売", "12_ミニ体験", "13_ミニ展示",
      "14_子供も楽しめる親子", "15_休憩可", "16_おひとり", "17_撮影禁止",
      "18_フラッシュ撮影禁止カメラ付き", "19_ビデオ撮影録音禁止", "20_接触禁止",
      "21_SNS禁止", "22_飲食禁止", "23_テイクアウト", "24_途中入退場可",
      "25_時間入替", "26_整理券", "27_先着順", "28_有料", "29_カンパ"]
  if array != nil
    count = 0
    while count < 19
      if tag6.size > 6
        break
      elsif array.include?(tag19[count])
        tag6.push(tagname[count] + ".ai")
      end
      count += 1
    end
  end
  return tag6[0], tag6[1], tag6[2], tag6[3], tag6[4], tag6[5]  #なしはnilで返す
end

###ここから###
starting = Process.clock_gettime(Process::CLOCK_MONOTONIC)  #処理時間の計測開始

id = Array.new  #企画配列の作成
promo = Array.new  #宣伝文配列作成
intro = Array.new  #学術紹介文配列作成
categ = Array.new  #カテゴリ配列作成
tag = Array.new  #タグ配列作成
icon = Array.new  #オリジナルアイコン配列作成

# 企画カテゴリファイルを読み込み、必要部分を配列化
index = 0  #カテゴリ配列変数を初期化
Dir.glob("オフィシャルパンフレット掲載企画カテゴリ*.csv").each do |item|  #ファイル名一覧の取得
  name = File.open(item)
  l_count = 0
  while true  #無限ループ
    line = name.gets  #1行読み込む
    if line == nil  #行がなければ抜ける
      break
    end
    line.chomp!  #改行を取り除く
    tuple = line.gsub(/\"/, '').split(',')  #1タプルごとのデータの配列化
    if l_count == 0
      length = tuple.size  #この表の要素数を取得
    elsif tuple[0] == ""
      break
    else
      categ[index] = Array.new
      categ[index][0] = tuple[length]  #idをカテゴリ配列[0]に保存
      categ[index][1] = tuple[1]  #大分類をカテゴリ配列[1]に保存
      if tuple[1].include?("参加型企画") || tuple[1].include?("ステージ")  #大分類
        tuple_data = tuple[1] + "‐" + tuple[2]
        categ[index][2] = tuple_data  #小分類をカテゴリ配列[2]に保存
      elsif tuple[1].include?("物品販売")  #大分類
        categ[index][2] = tuple[1] + "‐" + tuple[3]  #小分類をカテゴリ配列[2]に保存
      elsif tuple[1].include?("展示")  #大分類
        categ[index][2] = tuple[1] + "‐" + tuple[4]  #小分類をカテゴリ配列[2]に保存
      elsif tuple[1].include?("食品販売")  #大分類
        categ[index][2] = tuple[1] + "‐" + tuple[5]  #小分類をカテゴリ配列[2]に保存
      elsif tuple[1].include?("パフォーマンス")  #大分類
        categ[index][2] = tuple[1] + "‐" + tuple[6]  #小分類をカテゴリ配列[2]に保存
      elsif tuple[1].include?("その他")  #大分類
        categ[index][2] = tuple[1] + "‐" + tuple[7]  #小分類をカテゴリ配列[2]に保存
      end
      index += 1  #この企画の情報は終了
    end
    l_count += 1
  end  #無限ループここまで
end

# 宣伝文ファイルを読み込み、必要部分を配列化
index = 0
Dir.glob("オフィシャルパンフレット宣伝文*.csv").each do |item|  #ファイル名一覧の取得
  name = File.open(item)
  l_count = 0
  while true  #無限ループ
    line = name.gets  #1行読み込む
    if line == nil  #行がなければ抜ける
      break
    end
    line.chomp!  #改行を取り除く
    line.chop!  #末尾の「"」を取り除く
    tuple = line.split('","')  #1タプルごとのデータの配列化
    if l_count == 0
      length = tuple.size  #この表の要素数を取得
      tuple[0].slice!("\"")  #申請id保存
      regist_id = tuple[0].to_i
      if regist_id == 15
        break
        break
      end
    elsif tuple[0] == "\""
      break
    else
      promo[index] = Array.new
      promo[index][0] = tuple[length]  #idをカテゴリ配列[0]に保存
      if length == 4  #学術企画の場合
        if tuple[1] == "学術企画である"
          promo[index][1] = tuple[3]  #宣伝文を宣伝文配列[1]に保存
        else
          promo[index][1] = tuple[2]  #宣伝文を宣伝文配列[1]に保存
        end
      elsif length == 2
        promo[index][1] = tuple[1]  #宣伝文を宣伝文配列[1]に保存
      end
      index += 1  #この企画の情報は終了
    end
    l_count += 1
  end  #無限ループここまで
end

# 紹介文ファイルを読み込み、必要部分を配列化
index = 0
Dir.glob("オフィシャルパンフレット宣伝文*.csv").each do |item|  #ファイル名一覧の取得
  name = File.open(item)
  l_count = 0
  while true  #無限ループ
    line = name.gets  #1行読み込む
    if line == nil  #行がなければ抜ける
      break
    end
    line.chomp!  #改行を取り除く
    line.chop!  #末尾の「"」を取り除く
    tuple = line.split('","')  #1タプルごとのデータの配列化
    if l_count == 0
      length = tuple.size  #この表の要素数を取得
      tuple[0].slice!("\"")  #申請id保存
      regist_id = tuple[0].to_i
      if regist_id != 15  #申請idが15（学術紹介文）以外の場合
        break
        break
      end
    elsif tuple[0] == "\""
      break
    else
      intro[index] = Array.new
      intro[index][0] = tuple[length]  #紹介文を学術紹介文配列[0]に保存
      intro[index][1] = tuple[1]  #紹介文を学術紹介文配列[1]に保存
      index += 1  #この企画の情報は終了
    end
    l_count += 1
  end  #無限ループここまで
end

# タグアイコンファイルを読み込み、必要部分を配列化
index = 0
Dir.glob("オフィシャルパンフレットタグアイコン.csv").each do |item|  #ファイル名一覧の取得
  name = File.open(item)
  l_count = 0
  while true  #無限ループ
    line = name.gets  #1行読み込む
    if line == nil  #行がなければ抜ける
      break
    end
    line.chomp!
    line.chop!
    tuple = line.split('","')  #1タプルごとのデータの配列化
    if l_count == 0
      length = tuple.size  #この表の要素数を取得
    elsif tuple[0] == "\""
      next
    else
      tag[index] = Array.new
      tag[index][0] = tuple[length]  #idをタグ配列[0]に保存
      tagicon = tuple[1].split(",")
      tag[index] = tag[index] + tag_judge(tagicon)  #タグアイコンをタグ配列に保存
      index += 1  #この企画の情報は終了
    end
    l_count += 1
  end  #無限ループここまで
end
# オリジナルアイコンファイルを読み込み、必要部分を配列化
index = 0
Dir.glob("オフィシャルパンフレット学術オリジナルアイコン*.csv").each do |item|  #ファイル名一覧の取得
  name = File.open(item)
  l_count = 0
  while true  #無限ループ
    line = name.gets  #1行読み込む
    if line == nil  #行がなければ抜ける
      break
    end
    line.chomp!
    line.chop!
    tuple = line.split('","')  #1タプルごとのデータの配列化
    if l_count == 0
      length = tuple.size  #この表の要素数を取得
    elsif tuple[0] == "\""
      next
    else
      icon[index] = Array.new
      icon[index][0] = tuple[length]  #idをアイコン配列[0]に保存
      icon[index][1] = tuple[1]  #ファイル名をアイコン配列[1]に保存
      index += 1  #この企画の情報は終了
    end
    l_count += 1
  end  #無限ループここまで
end

# 全企画リストを読み込み、必要部分を配列化
index = 0
Dir.glob("全企画.csv").each do |item|  #ファイル名一覧の取得
  name = File.open(item)
  l_count = 0
  while true  #無限ループ
    line = name.gets  #1行読み込む
    if line == nil  #行がなければ抜ける
      break
    end
    line.slice!("\"")
    tuple = line.split('","')  #1タプルごとのデータの配列化
    if l_count == 0
    elsif tuple[1] == nil
      next
    elsif tuple[5].include?("\n")
      id[index] = Array.new
      id[index][0] = tuple[0]  #idを企画配列[0]に保存
      id[index][1] = tuple[1]  #企画団体名を企画配列[1]に保存
      id[index][2] = tuple[3]  #企画名を企画配列[2]に保存
      id[index][4] = tuple[2]  #企画団体名ヨミを企画配列[4]に保存
      id[index][5] = tuple[4]  #企画名ヨミを企画配列[5]に保存
      check = tuple[5]
      id[index][8] = tuple[5].chomp  #企画概要を企画配列[8]に保存
      while true
        if check.include?("\n")
          line = name.gets
          line.slice!("\"")
          tuple = line.split(',"')  #1タプルごとのデータの配列化
          check = tuple[0]
          id[index][8] += tuple[0].chomp
        else
          break
        end
      end
      id[index][3] = tuple[1].chop!  #参加区分を企画配列[3]に保存
      id[index][6] = tuple[2].chop!  #学術枠かを企画配列[6]に保存
      id[index][7] = tuple[3].chop!  #芸術枠かを企画配列[7]に保存
      index += 1  #この企画の情報は終了
    else
      id[index] = Array.new
      id[index][0] = tuple[0]  #idを企画配列[0]に保存
      id[index][1] = tuple[3]  #企画名を企画配列[1]に保存
      id[index][2] = tuple[1]  #企画団体名を企画配列[2]に保存
      id[index][3] = tuple[6]  #参加区分を企画配列[3]に保存
      id[index][4] = tuple[4]  #企画名ヨミを企画配列[4]に保存
      id[index][5] = tuple[2]  #企画団体名ヨミを企画配列[5]に保存
      id[index][6] = tuple[7]  #学術枠かを企画配列[6]に保存
      id[index][7] = tuple[8]  #芸術枠かを企画配列[7]に保存
      id[index][8] = tuple[5]  #企画概要を企画配列[8]に保存
      index += 1  #この企画の情報は終了
    end
    l_count += 1
  end  #無限ループここまで
end

data = Array.new  #最終整形配列
###配列データの結合ここから###
index = 0
while index < id.size  #企画数だけ繰り返す
  data[index] = Array.new
  num = 0
  while num < 8  #全企画情報を格納
    data[index][num] = id[index][num]
    num += 1
  end
  if data[index][3] == "ステージ" && data[index][6] == "〇"
    data[index][8] = "ステージ"
    data[index][9] = "ステージ‐学術"
  else
    count = 0
    while count < categ.size  #企画カテゴリの情報格納
      if categ[count][0] == data[index][0]  #idの一致
        data[index][8] = categ[count][1]  #大分類
        data[index][9] = categ[count][2]  #小分類
        break
      elsif count == categ.size - 1  #申請なし
        data[index][8] = ""  #大分類
        data[index][9] = ""  #小分類
      end
      count += 1
    end
  end
  count = 0
  while count < promo.size  #宣伝文の情報格納
    if promo[count][0] == data[index][0]  #idの一致
      data[index][10] = promo[count][1]  #宣伝文
      break
    elsif count == promo.size - 1  #申請なし
      data[index][10] = ""  #宣伝文
    end
    count += 1
  end
  count = 0
  while count < tag.size  #タグアイコン
    if tag[count][0] == data[index][0]
      num = 1
      while num < 7
        data[index][num + 10] = tag[count][num]
        num += 1
      end
    end
    count += 1
  end
  data[index][17] = id[index][8]  #企画概要
  count = 0
  while count < intro.size  #学術企画紹介文の情報格納
    if intro[count][0] == data[index][0]  #idの一致
      data[index][18] = intro[count][1]  #紹介文
      break
    elsif count == intro.size - 1  #申請なし
      data[index][18] = ""  #紹介文
    end
    count += 1
  end
  count = 0
  while count < icon.size  #学術オリジナルアイコンの情報格納
    if icon[count][0] == data[index][0]  #idの一致
      data[index][19] = icon[count][1]  #オリジナルアイコン
      break
    elsif count == intro.size - 1  #申請なし
      data[index][19] = ""  #オリジナルアイコン
    end
    count += 1
  end
  index += 1
end
###配列データの結合ここまで###

###データチェックここから###
print("---ERROR履歴------\n\n")
index = 0
while index < data.size
  data[index][20] = "N"
  #* 文字数制限チェック *#
  mojisu = count_char(data[index][10])  #宣伝文の文字数カウント
  if data[index][6] == "〇" #学術枠の場合
    if mojisu > 25
      print("    E01: 宣伝文超過*25+", mojisu-25, "字(id=", data[index][0], ")\n")
      data[index][20] += ";E01"  #宣伝文超過（学術）エラーコードE01
    end
    shokai = count_char(data[index][18])  #紹介文の文字数カウント
    if shokai > 105
      print("    E03: 紹介文超過*105+", shokai-105, "字(id=", data[index][0], ")\n")
      data[index][20] += ";E03"  #紹介文超過（学術）エラーコードE03
    end
    data[index][20] += emoji_opt(data[index][18], data[index][0], true)  #文字化けチェック
  else  #非学術枠の場合
    if mojisu > 30
      print("    E02: 宣伝文超過*30+", mojisu-30, "字(id=", data[index][0], ")\n")
      data[index][20] += ";E02"  #宣伝文超過エラーコードE02
    end
  end
  #*文字化けチェック*#
  data[index][20] += emoji_opt(data[index][10], data[index][0], false)
  index += 1  #この企画の情報は終了
end
###データチェックここまで###

# csv形式でpanf-out.csvに出力
File.open("panf-out.csv", "w") do |out|
  # csv形式で1行目出力
  out.print("id,\"企画名\",\"企画団体名\",\"参加区分\",\"企画名ヨミ\",")
  out.print("\"企画団体名ヨミ\",\"学術枠\",\"芸術枠\",\"大分類\",\"小分類\",\"宣伝文\",")
  out.print("\"tag1\",\"tag2\",\"tag3\",\"tag4\",\"tag5\",\"tag6\",\"企画概要\",")
  out.print("\"学術紹介\",\"学術オリジナルアイコン\",\"ERROR\"\n")
  # csv形式に２行目以降出力
  index = 0  #繰り返し変数初期化
  while index < data.size  #繰り返す
    out.print(data[index][0].to_i)
    count = 1  #繰り返し変数の初期化
    while count < data[index].size  #全要素書き出す
      out.print(",\"", data[index][count], "\"")
      count += 1  #繰り返し変数更新
    end
    out.print("\n")  #区切りの改行
    index += 1  #繰り返し変数更新
  end
end

print("\n\n基本情報を統合し、panf-out.csvに書き出しました。\n")
print("引き続き実施場所情報を統合します...\n\n")


#配置情報を統合し、ページを割り振るプログラム
###メソッド定義###
#エリア番号を付与するメソッド
# 引数：エリア名
# 返り値：エリア番号
def areaindex(areaname, tent)
  areaname = areaname.to_s
  tent = tent.to_s
  if areaname.include?("松美池周辺")
    if tent.include?("C") || tent.include?("D")
      areano = "23"
    elsif tent.include?("E") || tent.include?("F")
      areano = "25"
    else
      areano = "25"
    end
  elsif areaname.include?("図書館周辺")
    if tent.include?("C") || tent.include?("D")
      areano = "03"
    elsif tent.include?("E")
      areano = "05"
    elsif tent.include?("F")
      areano = "07"
    else
      areano = "07"
    end
  elsif areaname.include?("大学会館周辺")
    areano = "31"
  elsif areaname.include?("5C棟周辺")
    if tent.include?("A") || tent.include?("B")
      areano = "35"
    else
      areano = "37"
    end
  elsif areaname.include?("2A棟・3A棟周辺")
    areano = "01"
  elsif areaname.include?("1A棟・1C棟・1D棟周辺")
    if tent.include?("A") || tent.include?("B")
      areano = "19"
    else
      areano = "21"
    end
  else
    areano = "43"  #エリア未定/不明
  end
  return areano
end

#エリア番号を付与するメソッド
# 引数：教室
# 返り値：エリア番号
def areaindex_okunai(area, room)
  area = area.to_s
  if area.include?("第３エリア")
    areano = "09"
  elsif area.include?("第２エリア")
    areano = "13"
  elsif area.include?("第１エリア")
    areano = "27"
  elsif area.include?("会館エリア")
    areano = "33"
  elsif area.include?("体")
    if /5[A-Z][4-7]+/ =~ room
      areano = "39"
    elsif /5[A-Z]+/ =~ room
      areano = "40"
    elsif /6[A-Z][4-7]+/ =~ room
      areano = "41"
    else
      areano = "42"
    end
  elsif area.include?("その他")
    if room.include?("中央図書館")
      areano = "09"
    elsif room.include?("総合研究")
      areano = "09"
    elsif room.include?("開学記念")
      areano = "42"
    else
      areano = "43"
    end
  else
    areano = "44"  #エリア未定/不明/各所
  end
  return areano
end

#テントソートコードを付与するメソッド
# 引数：テント番号
# 返り値：ソート用コード, テント番号
def tentindex(tent)
  str = tent.to_s
  chr = str.slice!(0)
  code = chr.to_s.tr("A-I", "1-9")
  str = sprintf("%02d", str.to_i).to_s
  code = code + str
  return code.to_s
end

#企画分類アイコンのファイル名を取得するメソッド
# 引数：小分類
# 返り値：企画分類アイコンのファイル名
def kikakuicon_add(category)
  if category == nil
    name = nil
  elsif /^参加型企画+/ =~ category
    if category.include?("会")  #座談会・相談会はカウンターアイコン
      name = "counter.ai"
    else  #それ以外の参加型は参加アイコン
      name = "joint.ai"
    end
  elsif /^物品販売+/ =~ category
    name = "present.ai"
  elsif /^展示‐芸+/ =~ category
    name = "picture.ai"
  elsif /^展示‐学+/ =~ category
    name = "pencil.ai"
  elsif /^展示‐+/ =~ category
    name = "projector.ai"
  elsif /^食品販売+/ =~ category
    if category.include?("スイーツ")  #スイーツはケーキアイコン
      name = "cake.ai"
    elsif category.include?("ドリンク") || category.include?("喫茶")
      name = "drink.ai"  #飲料系はコップアイコン
    else  #それ以外の食品販売はレストランアイコン
      name = "restaurant.ai"
    end
  elsif /^パフォーマンス+/ =~ category
    name = "microphone.ai"
  elsif /^ステージ+/ =~ category
    name = "stage.ai"
  else
    name = "etc.ai"
  end
  return name
end

#学術企画ページの掲載順序を付与するメソッド
# 引数：id, 学術ハッシュ
# 返り値：学術企画参照ページ番号, 学術企画参照番地
def tsukulab(id, lab)
  if lab.size > 0 && lab.key?(id.to_s)
    page, address = lab[id.to_s]
  else
    page = (lab.size).div(7) + 1
    address = (lab.size) % 7 + 1
    store_p = [page, address]
    lab.store(id.to_s, store_p)
  end
  return page, address
end

#ワールドクイズラリーアイコンを付与するメソッド
# 引数：id, ワールドクイズラリー企画配列
# 返り値：ファイル名
def jitsuiadd(kikakuid, wqlist)
  result = ""
  i = 0
  while i < wqlist.size
    if kikakuid.to_s == wqlist[i][0].to_s
      result = "ワールドクイズラリー.ai"
    end
    i += 1
  end
  return result
end

###ここから###
require 'csv'
id = Array.new  #企画配列の作成
okugai = Array.new  #屋外配列作成
okunai = Array.new  #屋内配列作成
icon = Array.new  #アイコン配列作成
ord = Array.new  #一般配列作成
stg = Array.new  #ステージ配列作成
lab = Hash.new  #学術ハッシュ作成
wqlist = Array.new  #ワールドクイズラリー企画配列作成

# 企画情報（屋外）のCSVを読み込み、必要部分を配列化
index = 0
filename = Dir.glob("*企画情報*屋外*.csv").max_by { |fn| File.birthtime(fn) }
if filename == nil
  print("屋外の企画情報のCSVファイルが見つかりませんでした。\n同じフォルダに企画情報のCSVファイルを置いてから実行してください。\n")
  exit
end
CSV.foreach(filename) do |tuple|  #ファイル名一覧の取得
  if tuple[0] == "id"
    next
  elsif tuple[0] == nil
    break
  elsif tuple[1].to_s.include?("(企画辞退)")
    next
  else
    okugai[index] = Array.new
    okugai[index][0] = tuple[0].to_i  #idを屋外配列[0]に保存
    okugai[index][1] = tuple[4]  #エリアを屋外配列[1]に保存
    if tuple[5].to_s.include?("/")
      multiarea = tuple[4].split("/")
      multitent = tuple[5].split("/")
      k = 0
      while k < multiarea.size
        if k != 0
          okugai[index + k] = Array.new
          okugai[index + k][0] = tuple[0].to_i  #idを屋外配列[0]に保存
          okugai[index + k][1] = tuple[4]  #エリアを屋外配列[1]に保存
        end
        okugai[index + k][2] = areaindex(multiarea[k], multitent[k])  #エリアの順番付け
        okugai[index + k][4] = sprintf("%3s", multitent[k]).strip  #テント番号を屋外配列[4]に保存
        okugai[index + k][3] = tentindex(multitent[k])  #テントIDを屋外配列[3]に保存
        okugai[index + k][5] = tuple[6]  #2日を屋外配列[5]に保存
        okugai[index + k][6] = tuple[7]  #3日を屋外配列[6]に保存
        okugai[index + k][7] = tuple[8]  #4日を屋外配列[7]に保存
        okugai[index + k][8] = tuple[11]  #テント外企画を屋外配列[8]に保存
        okugai[index + k][9] = ";E05"  #移動のエラーコードを記録
        if okugai[index + k][3] == 0  #テント番号空値の場合
          okugai[index + k][9] += ";E07"  #テント番号不明のエラーコードを記録
        end
        if okugai[index][2] == 95  #エリア不明の場合
          okugai[index][9] += ";E08"  #エリア不明のエラーコードを記録
        end
        k += 1
      end
      index = index + k  #この企画の配置情報は終了
    else
      okugai[index][2] = areaindex(tuple[4], tuple[5])  #エリアの順番付け
      okugai[index][4] = sprintf("%3s", tuple[5]).strip   #テント番号を屋外配列[4]に保存
      okugai[index][3] = tentindex(tuple[5])  #テントIDを屋外配列[3]に保存
      okugai[index][5] = tuple[6]  #2日を屋外配列[5]に保存
      okugai[index][6] = tuple[7]  #3日を屋外配列[6]に保存
      okugai[index][7] = tuple[8]  #4日を屋外配列[7]に保存
      okugai[index][8] = tuple[11]  #テント外企画を屋外配列[8]に保存
      okugai[index][9] = ""  #ERRORなし
      if okugai[index][3] == 0  #テント番号空値の場合
        okugai[index][9] += ";E07"  #テント番号不明のエラーコードを記録
      end
      if okugai[index][2] == 95  #エリア不明の場合
        okugai[index][9] += ";E08"  #エリア不明のエラーコードを記録
      end
      index += 1  #この企画の配置情報は終了
    end
  end
end  #CSV読み込みここまで

# 企画情報（屋内）のCSVを読み込み、必要部分を配列化
index = 0
filename = Dir.glob("*企画情報*屋内*.csv").max_by { |fn| File.birthtime(fn) }
if filename == nil
  print("屋内の企画情報のCSVファイルが見つかりませんでした。\n同じフォルダに企画情報のCSVファイルを置いてから実行してください。\n")
  exit
end
CSV.foreach(filename) do |item|  #ファイル名一覧の取得
  if item[0] == "id"
    next
   elsif item[0] == nil
    break
  elsif item[2].to_s.include?("企画辞退") || item[2].to_s.include?("企画中止") || item[2].to_s.include?("企画自体")
    next
  else
    okunai[index] = Array.new
    okunai[index][0] = item[0].to_i  #idを屋内配列[0]に保存
    okunai[index][1] = item[2]  #エリアを屋内配列[1]に保存
    if item[3].to_s.include?(",")  #複数教室の場合
      if item[3].to_s.include?(":")  #日により企画場所変更の場合
        multiroom = item[3].split(",")
        k = 0
        while k < multiroom.size
          if k != 0
            okunai[index + k] = Array.new
            okunai[index + k][0] = item[0].to_i  #idを屋内配列[0]に保存
            okunai[index + k][1] = item[2]  #エリアを屋内配列[1]に保存
          end
          okunai[index + k][2] = areaindex_okunai(item[2], multiroom[k].sub(/...:/, "").strip)  #エリアの順番付け
          okunai[index + k][3] = multiroom[k].sub(/...:/, "").strip  #教室を屋内配列[3]に保存
          okunai[index + k][4] = multiroom[k].sub(/...:/, "").strip  #教室を屋内配列[4]に保存
          okunai[index + k][5] = "×"  #2日を屋外配列[5]に保存
          okunai[index + k][6] = "×"  #3日を屋外配列[6]に保存
          okunai[index + k][7] = "×"  #4日を屋外配列[7]に保存
          if /前夜祭+/ =~ multiroom[k]
            okunai[index + k][5] = item[4]  #2日を屋外配列[5]に保存
          end
          if /1日目+/ =~ multiroom[k]
            okunai[index + k][6] = item[5]  #3日を屋外配列[6]に保存
          end
          if /2日目+/ =~ multiroom[k]
            okunai[index + k][7] = item[6]  #4日を屋外配列[7]に保存
          end
          okunai[index + k][8] = item[7]  #テント外企画を屋外配列[8]に保存
          okunai[index + k][9] = ";E05"  #移動のエラーコードを記録
          if okugai[index][2] == 99  #エリア不明の場合
            okugai[index][9] += ";E08"  #エリア不明のエラーコードを記録
          end
          k += 1
        end
        index = index + k  #この企画の配置情報は終了
      else
        room_m = item[3].split(",")
        item[3] = room_m[0]
        j = 1
        while j < room_m.size
          if /\d[A-Z]+/ =~ room_m[j].to_s
            item[3] = item[3] + "/" + room_m[j]
          else
            item[3] = item[3] + "," + room_m[j]
          end
          j += 1
        end
        multiroom = item[3].split("/")
        k = 0
        while k < multiroom.size
          if k != 0
            okunai[index + k] = Array.new
            okunai[index + k][0] = item[0].to_i  #idを屋内配列[0]に保存
            okunai[index + k][1] = item[2]  #エリアを屋内配列[1]に保存
          end
          okunai[index + k][2] = areaindex_okunai(item[2], multiroom[k])  #エリアの順番付け
          okunai[index + k][3] = multiroom[k]  #教室を屋内配列[3]に保存
          okunai[index + k][4] = item[3]  #教室を屋内配列[4]に保存
          okunai[index + k][5] = item[4]  #2日を屋外配列[5]に保存
          okunai[index + k][6] = item[5]  #3日を屋外配列[6]に保存
          okunai[index + k][7] = item[6]  #4日を屋外配列[7]に保存
          okunai[index + k][8] = item[7]  #テント外企画を屋外配列[8]に保存
          okunai[index + k][9] = ";E06"  #複数場所のエラーコードを記録
          if okunai[index + k][2] == 99  #エリア不明の場合
            okunai[index + k][9] += ";E08"  #エリア不明のエラーコードを記録
          end
          k += 1
        end
        index = index + k  #この企画の配置情報は終了
      end
    else
      okunai[index][2] = areaindex_okunai(item[2], item[3])  #エリアの順番付け
      okunai[index][3] = item[3]  #教室を屋内配列[3]に保存
      okunai[index][4] = item[3]  #教室を屋内配列[4]に保存
      i = 5
      while i < 9
        okunai[index][i] = item[i - 1]  #屋内配列に保存
        i += 1
      end
      okunai[index][9] = ""  #ERRORなし
      if okunai[index][2] == 99  #エリア不明の場合
        okunai[index][9] += ";E08"  #エリア不明のエラーコードを記録
      end
      index += 1  #この企画の配置情報は終了
    end
  end
end  #CSV読み込みここまで

# ステージタイムテーブルのCSVを読み込み、必要部分を配列化
index = 0
Dir.glob("*TT*.csv").each do |filename|
  if filename == nil
    print("ステージタイムテーブルのCSVファイルが見つかりませんでした。\n同じフォルダにステージタイムテーブルのCSVファイルを置いてから実行してください。\n")
    exit
  end
  date = 2  #日の初期値
  place = ""  #場所の初期値
  CSV.foreach(filename) do |tuple|  #ファイル名一覧の取得
    if tuple[0] != nil && tuple[1] == nil
      name = filename.split("_")
      File.open("panf-note.txt", "w") do |note|
        note.print("ステージ:", name[0], ",", date, "日,内容:", tuple[0], "\n")
      end
      print("\n注記をpanf-note.txtに書き出しました。\n\n")
    elsif tuple[0].to_s.include?("前夜祭")
      date = 2
    elsif tuple[0].to_s.include?("1日目")
      date = 3
      if tuple[0].to_s.include?("ホール")
        place = "（ホール）"
      elsif tuple[0].to_s.include?("講堂")
        place = "（講堂）"
      end
    elsif tuple[0].to_s.include?("2日目")
      date = 4
      if tuple[0].to_s.include?("ホール")
        place = "（ホール）"
      elsif tuple[0].to_s.include?("講堂")
        place = "（講堂）"
      end
    elsif tuple[0] == nil
      next
    else
      name = filename.split("_")
      stg[index] = Array.new
      stg[index][0] = tuple[0].to_s  #企画名をステージ配列[0]に保存
      stg[index][1] = date  #日をステージ配列[1]に保存
      stg[index][2] = name[0] + place  #場所をステージ配列[2]に保存
      stg[index][3] = tuple[1]  #開始時刻をステージ配列[3]に保存
      stg[index][4] = tuple[2]  #終了時刻をステージ配列[4]に保存
      stg[index][5] = sprintf("%02d", index)  #順序をステージ配列[6]に保存
      stg[index][6] = tuple[3]  #idをステージ配列[6]に保存
      index += 1  #この企画の配置情報は終了
    end
  end  #CSV読み込みここまで
end

# ワールドクイズラリー企画リストを読み込み、必要部分を配列化
index = 0
Dir.glob("*ワールドクイズラリー*.txt").each do |item|  #ファイル名一覧の取得
  CSV.foreach(item,:col_sep => "\t") do |row|
    wqlist[index] = Array.new
    wqlist[index] = row
    index += 1
  end
end

# panf-out.csvを読み込み、必要部分を配列化
index = 0
Dir.glob("panf-out*.csv").each do |item|  #ファイル名一覧の取得
  name = File.open(item)
  line = name.gets  #1行読み込む
  while true  #無限ループ
    line = name.gets  #1行読み込む
    if line == nil  #行がなければ抜ける
      break
    end
    line.chomp!
    tuple = line.split(',"')  #1タプルごとのデータの配列化
    if /\A\(企画\s+/ =~ tuple[1]  #企画辞退をとばす
      next
    else
      id[index] = Array.new
      id[index][0] = tuple[0].to_i  #idを企画配列[0]に保存
      i = 1
      while i < 21
        id[index][i] = tuple[i].chop  #企画配列に保存
        i += 1
      end
      if id[index][10] == ""
        id[index][20] += ";E10"  #宣伝文空値のエラーコードを記録
      end
      if id[index][9] == ""
        id[index][20] += ";E09"  #企画カテゴリ空値のエラーコードを記録
      end
      if id[index][6] == "〇"
        if id[index][18] == ""
          id[index][20] += ";E11"  #紹介文空値のエラーコードを記録
        end
        if id[index][19] == ""
          id[index][20] += ";E12"  #学術オリジナルアイコンファイル名空値のエラーコードを記録
        end
      end
      index += 1  #この企画の情報は終了
    end
  end  #無限ループここまで
end
if id.size == 0
  print("panf-out.csvが見つかりませんでした。\ncomb.rbを実行して、同じフォルダにpanf-out.csvを置いてから実行してください。\n")
  exit
end

###ページ順にソート###
ord = okugai + okunai  #一般企画配列を作成
#場所順に並び替え(ページ順=>ソートコード=>前夜祭=>1日目)
ord.sort! {|a, b| (a[2].to_i <=> b[2].to_i).nonzero? || (a[3].to_s <=> b[3].to_s).nonzero? || (b[5].to_s <=> a[5].to_s).nonzero? || (b[6].to_s <=> a[6].to_s)}
stg.sort! {|a, b| (a[1] <=> b[1]).nonzero? || (a[2] <=> b[2]).nonzero? || (a[5] <=> b[5])}

###一般企画の配列データの結合###
index = 0
while index < ord.size  #企画数だけ繰り返す
  error = ord[index][9].to_s  #エラーコードの記録
  count = 0
  while count < id.size
    if ord[index][0] == id[count][0]  #結合
      num = 1
      while num < 21
        ord[index][num + 8] = id[count][num]
        num += 1
      end
      break
    end
    count += 1
  end
  if count == id.size  #もし企画情報が見つからなかった場合
    ord[index][9] = ""
    ord[index][28] = error + ";E33"  #エラーコードを返す
  else
    ord[index][28] = ord[index][28].to_s + error  #エラーコードのデータ統合
    ord[index][29] = kikakuicon_add(ord[index][17])  #企画分類アイコンのファイル名を格納
  end
  if ord[index][14] == "〇"  #学術枠ならば、つくラボ参照ページを格納
    ord[index][30], ord[index][31] = tsukulab(ord[index][0], lab)
    ord[index][32] = "01_Tsukulab_mini.ai"  #学術枠アイコンのファイル名を格納
  end
  if ord[index][15] == "〇"  #芸術枠ならば、芸術アイコンのファイル名を格納
    if ord[index][32] == nil
      ord[index][32] = "04_geijutsuwaku.ai"
      ord[index][33] = jitsuiadd(ord[index][0], wqlist)  #予備学実委アイコンのファイル名格納
    else
      ord[index][33] = "04_geijutsuwaku.ai"
      ord[index][35] = jitsuiadd(ord[index][0], wqlist)  #予備学実委アイコンのファイル名格納
    end
  end
  if ord[index][32] == nil  #まだ何も学実委アイコンが付与されていない場合
    ord[index][32] = jitsuiadd(ord[index][0], wqlist)  #予備学実委アイコンのファイル名格納
  end
  ord[index][34] = sprintf("%03d", index + 1)  #一般シリアルナンバーの付与(主キー)
  index += 1  #この企画の処理は終了する
end

###ステージ企画の配列データの結合###
index = 0
while index < stg.size  #企画数だけ繰り返す
  count = 0
  while count < id.size
    if id[count][3] == "ステージ" && (stg[index][0] == id[count][1] || stg[index][6] == id[count][0].to_s)  #結合
      stg[index][6] = id[count][20].to_s  #エラーコードの統合
      num = 0
      while num < 20
        stg[index][num + 7] = id[count][num]
        num += 1
      end
      break
    end
    count += 1
  end
  if stg[index][6].to_s.include?(";") == false
    count = 0
    while count < id.size
      if stg[index][0] == id[count][2]  #団体名で結合を試みる
        stg[index][6] = id[count][20].to_s  #エラーコードの統合
        stg[index][7] = id[count][0]  #idの統合
        stg[index][8] = id[count][2]  #企画名に基本情報の団体名を入れる
        stg[index][9] = id[count][1]  #企画団体名に基本情報の企画名を入れる
        num = 3
        while num < 20
          stg[index][num + 7] = id[count][num]
          num += 1
        end
        stg[index][6] = stg[index][6].to_s + ";E31"
        break
      end
      count += 1
    end
    if stg[index][6].to_s.include?(";") == false  #企画情報紐づけ失敗のエラーコードを返す
      stg[index][6] = "N;E30"
    end
  end
  if stg[index][13].to_s == "〇"  #学術枠ならば、つくラボ参照ページを格納
    stg[index][27], stg[index][28] = tsukulab(stg[index][7], lab)
    stg[index][29] = "01_Tsukulab_mini.ai"  #学術枠アイコンのファイル名を格納
  end
  index += 1
end

###一般企画配列をcsv形式でpanf-ord.csvに出力###
File.open("panf-ord.csv", "w") do |out|
  # csv形式で1行目出力
  out.print("ORDｼﾘｱﾙ,id,\"場所\",\"ページ順序\",\"ソートコード\",\"テント番号/屋内教室\",")
  out.print("\"2日\",\"3日\",\"4日\",\"テント外企画\",\"企画名\",\"企画団体名\",")
  out.print("\"参加区分\",\"企画名ヨミ\",\"企画団体名ヨミ\",\"学術枠\",\"芸術枠\",")
  out.print("\"大分類\",\"小分類\",\"宣伝文\",\"学実委icon1\",\"学実委icon2\",\"学実委icon3\",\"tag1\",\"tag2\",\"tag3\",\"tag4\",")
  out.print("\"tag5\",\"tag6\",\"企画概要\",\"学術紹介\",\"学術オリジナルアイコン\",\"学術参照ページ\",\"学術参照番地\",\"企画分類アイコン\",\"ERROR\"\n")
  # csv形式に２行目以降出力
  index = 0  #繰り返し変数初期化
  while index < ord.size  #繰り返す
    out.print(ord[index][34])  #一般シリアルナンバーを第1列に出力
    count = 0  #繰り返し変数の初期化
    while count < 19  #全要素書き出す
      out.print(",\"", ord[index][count], "\"")
      count += 1  #繰り返し変数更新
    end
    out.print(",\"", ord[index][32], "\",\"", ord[index][33], "\",\"", ord[index][35], "\"")  #学実委アイコンファイル名の出力
    while count < 28  #全要素書き出す
      out.print(",\"", ord[index][count], "\"")
      count += 1  #繰り返し変数更新
    end
    out.print(",\"", ord[index][30], "\",\"", ord[index][31], "\"")  #学術ページ番号の出力
    out.print(",\"", ord[index][29], "\"")  #企画分類アイコン属性の出力
    out.print(",\"", ord[index][28], "\"")  #エラーコード属性の出力
    out.print("\n")  #区切りの改行
    index += 1  #繰り返し変数更新
  end
end

###ステージ企画配列をcsv形式でpanf-stg.csvに出力###
File.open("panf-stg.csv", "w") do |out|
  # csv形式で1行目出力
  out.print("STGｼﾘｱﾙ,id,\"日付\",\"場所\",\"開始時間\",\"終了時間\",\"企画名\",\"企画団体名\",")
  out.print("\"参加区分\",\"企画名ヨミ\",\"企画団体名ヨミ\",\"学術枠\",\"芸術枠\",")
  out.print("\"大分類\",\"小分類\",\"宣伝文\",\"学実委icon\",\"企画概要\",\"学術紹介\",\"学術オリジナルアイコン\",\"学術参照ページ\",\"学術参照番地\",\"ERROR\"\n")
  # csv形式に２行目以降出力
  index = 0  #繰り返し変数初期化
  while index < stg.size  #繰り返す
    out.print(stg[index][5])  #ステージシリアルナンバーを第1列に出力
    out.print(",\"", stg[index][7], "\"")  #idを第2列に出力
    count = 1  #繰り返し変数の初期化
    while count < 5  #全要素書き出す
      out.print(",\"", stg[index][count], "\"")
      count += 1  #繰り返し変数更新
    end
    out.print(",\"", stg[index][0], "\"")  #企画名を出力
    count = 9  #繰り返し変数の初期化
    while count < 18  #全要素書き出す
      out.print(",\"", stg[index][count], "\"")
      count += 1  #繰り返し変数更新
    end
    out.print(",\"", stg[index][29], "\"")  #学術アイコンファイル名を出力
    out.print(",\"", stg[index][24], "\"")  #企画概要を出力
    out.print(",\"", stg[index][25], "\"")  #学術紹介を出力
    out.print(",\"", stg[index][26], "\"")  #オリジナルアイコンを出力
    out.print(",\"", stg[index][27], "\",\"", stg[index][28], "\"")  #学術参照ページを出力
    out.print(",\"", stg[index][6], "\"")  #エラーコードを出力
    out.print("\n")  #区切りの改行
    index += 1  #繰り返し変数更新
  end
end

ending = Process.clock_gettime(Process::CLOCK_MONOTONIC)  #処理時間の計測終了
print("\n統合データをpanf-ord.csvとpanf-stg.csvに書き出しました。\n")
print("処理開始からの経過時間: ", ending - starting, "秒\n(Enterを押して終了します)\n")
ok = gets.chomp!
exit
