######  月毎の住宅着工のエクセルシートから四半期データを作成

library(tidyverse)
library(readxl)

# 作業フォルダを指定、適宜修正。
# この下にe-statから入手した下記のe-statの住宅着工統計のデータを置く
# 「統計データをみる」から「分野からさがす」「建築着工統計調査」「住宅着工統計」「月次」
#  適宜の年月を選択し、表番号17のエクセルファイルをダウンロードして、下記フォルダの下におく
#  同様に「～b017.xls」ファイルを順次ダウンロードして保存。
#  ただし、最も古い年月は四半期の最初(1,4,7,10月)とすること。

setwd("F:/jinko_kenchiku/jutakuchakko")

fnames <- dir(pattern="^\\d{4}b017.xls")    #各月の名の取得　正規表現を使用
fnames

sheet_n = c(1,6,7)       # sheet１枚目に大阪府、6,7枚目に大阪府各市町村がある
city <- vector()         # 空vectorを作成
時点 <- vector()
分類 <- vector()
戸数 <- vector()

# 必要なデータを抜き出す　-----エクセルのデータ形式にあわせて-----
    #   時間かかる・・・ 
for ( i in seq_along(fnames) ) {
  for (j in seq_along(sheet_n) ) {
    e_df <- as.data.frame( read_excel(fnames[[i]],sheet = sheet_n[[j]]) )
    os <- which(str_sub(e_df[,1],1,2) == "27")   # 1列目の頭2文字が27(大阪)の行番号を取得
    for ( k in seq_along(os)) {
      for ( m in 1:5) {     # 5分類(合計・持家・貸家・給与住宅・分譲住宅)をよみこみ
        city <- c( city, e_df[os[k],1]) 
        時点 <- c( 時点, as.character(as.integer(str_sub(fnames[i],1,4)) + 198800 ) )
        分類 <- c( 分類, e_df[os[k]+m,1] )
        戸数 <- c( 戸数, e_df[os[k]+m,3] )
      }
    }
  }
}
ken <- data.frame(city,時点,分類,戸数)
# 月ごとの表を作成
ken2 <- spread(ken,key=時点, value=戸数 )
# 戸数をファクターから数に
ken2[3:length(ken2)] <- map(ken2[3:length(ken2)], as.character)
ken2[3:length(ken2)] <- map(ken2[3:length(ken2)], as.integer)

#########  四半期のデータを作成
ken_q <- ken2[1:2]
ken_q$市区町村名 <- str_sub(ken_q$city,6,15)
ken_q[1:5, 3] <- "大阪府"

# 大阪市と堺市を区名につけ加える
for ( i in 1:nrow(ken_q) ) {
  n <- as.integer( str_sub(ken_q[i,1],1,5) )
  if( 27100 < n & n < 27140 ) {
    ken_q[i,3] <- str_c("大阪市", ken_q[i,3])
  } else if ( 27140 < n & n < 27150 ) { 
    ken_q[i,3] <- str_c("堺市", ken_q[i,3])
  }
}

### 四半期データの作成
for (i in seq_along(ken2)) {
  if( i %% 3 == 0) {
    j <- ifelse(i+3 < length(ken2), i+2, length(ken2))
    q <- rowSums(ken2[i:j], na.rm=TRUE)
    ken_q <- cbind(ken_q, q)
    # 列名の作成
    if (i+2 ==j) {
      n <- str_c(str_sub(names(ken2[i]),1,4),"Q",
                 as.integer(str_sub(names(ken2[i]),5,6)) %/% 3 + 1)
    } else if (i+1==j ) {
      n <- str_c(names(ken2[i]), "-", str_sub(names(ken2[i+1]),5,6))
    } else {
      n <- names(ken2[i])
    }
    names(ken_q) <- c(names(ken_q[1:length(ken_q)-1]), n)
  }
}

###   グラフ作成
names(ken_q)
# エクセル的な横長の表はgatherしないと、ggplotではグラフにできない。
  # 数字で始まる列名は逆クォート``でかこう必要あり。 
ken_q_gather <- gather(ken_q,`2014Q1`:`2018Q2`,key="year_Q", value="戸数") 
ken_q_gather <- ken_q_gather %>% subset(分類 != "合計")　   # 分類が合計のものを外す。
# 堺市の区毎に
ken_q_gather %>% subset(str_detect(市区町村名,"堺市.+区")) %>% ggplot() + 
  geom_col(aes(year_Q, 戸数, fill=分類)) + facet_wrap(市区町村名~.) +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

### 持家、貸家、分譲住宅毎のcsvファイルを作成・保存 (エクセル用)
len <- length(ken_q)
mochi <- subset(ken_q, 分類=="持家")
kashi <- subset(ken_q, 分類=="貸家")
bunjo <- subset(ken_q, 分類=="分譲住宅")  
write.csv(mochi, "mochi.csv")
write.csv(kashi, "kashi.csv")
write.csv(bunjo, "bunjo.csv")