######  �����̏Z��H�̃G�N�Z���V�[�g����l�����f�[�^���쐬

library(tidyverse)
library(readxl)

# ��ƃt�H���_���w��A�K�X�C���B
# ���̉���e-stat������肵�����L��e-stat�̏Z��H���v�̃f�[�^��u��
# �u���v�f�[�^���݂�v����u���삩�炳�����v�u���z���H���v�����v�u�Z��H���v�v�u�����v
#  �K�X�̔N����I�����A�\�ԍ�17�̃G�N�Z���t�@�C�����_�E�����[�h���āA���L�t�H���_�̉��ɂ���
#  ���l�Ɂu�`b017.xls�v�t�@�C���������_�E�����[�h���ĕۑ��B
#  �������A�ł��Â��N���͎l�����̍ŏ�(1,4,7,10��)�Ƃ��邱�ƁB

setwd("F:/jinko_kenchiku/jutakuchakko")

fnames <- dir(pattern="^\\d{4}b017.xls")    #�e���̖��̎擾�@���K�\�����g�p
fnames

sheet_n = c(1,6,7)       # sheet�P���ڂɑ��{�A6,7���ڂɑ��{�e�s����������
city <- vector()         # ��vector���쐬
���_ <- vector()
���� <- vector()
�ː� <- vector()

# �K�v�ȃf�[�^�𔲂��o���@-----�G�N�Z���̃f�[�^�`���ɂ��킹��-----
    #   ���Ԃ�����E�E�E 
for ( i in seq_along(fnames) ) {
  for (j in seq_along(sheet_n) ) {
    e_df <- as.data.frame( read_excel(fnames[[i]],sheet = sheet_n[[j]]) )
    os <- which(str_sub(e_df[,1],1,2) == "27")   # 1��ڂ̓�2������27(���)�̍s�ԍ����擾
    for ( k in seq_along(os)) {
      for ( m in 1:5) {     # 5����(���v�E���ƁE�݉ƁE���^�Z��E�����Z��)����݂���
        city <- c( city, e_df[os[k],1]) 
        ���_ <- c( ���_, as.character(as.integer(str_sub(fnames[i],1,4)) + 198800 ) )
        ���� <- c( ����, e_df[os[k]+m,1] )
        �ː� <- c( �ː�, e_df[os[k]+m,3] )
      }
    }
  }
}
ken <- data.frame(city,���_,����,�ː�)
# �����Ƃ̕\���쐬
ken2 <- spread(ken,key=���_, value=�ː� )
# �ː����t�@�N�^�[���琔��
ken2[3:length(ken2)] <- map(ken2[3:length(ken2)], as.character)
ken2[3:length(ken2)] <- map(ken2[3:length(ken2)], as.integer)

#########  �l�����̃f�[�^���쐬
ken_q <- ken2[1:2]
ken_q$�s�撬���� <- str_sub(ken_q$city,6,15)
ken_q[1:5, 3] <- "���{"

# ���s�ƍ�s���於�ɂ�������
for ( i in 1:nrow(ken_q) ) {
  n <- as.integer( str_sub(ken_q[i,1],1,5) )
  if( 27100 < n & n < 27140 ) {
    ken_q[i,3] <- str_c("���s", ken_q[i,3])
  } else if ( 27140 < n & n < 27150 ) { 
    ken_q[i,3] <- str_c("��s", ken_q[i,3])
  }
}

### �l�����f�[�^�̍쐬
for (i in seq_along(ken2)) {
  if( i %% 3 == 0) {
    j <- ifelse(i+3 < length(ken2), i+2, length(ken2))
    q <- rowSums(ken2[i:j], na.rm=TRUE)
    ken_q <- cbind(ken_q, q)
    # �񖼂̍쐬
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

###   �O���t�쐬
names(ken_q)
# �G�N�Z���I�ȉ����̕\��gather���Ȃ��ƁAggplot�ł̓O���t�ɂł��Ȃ��B
  # �����Ŏn�܂�񖼂͋t�N�H�[�g``�ł������K�v����B 
ken_q_gather <- gather(ken_q,`2014Q1`:`2018Q2`,key="year_Q", value="�ː�") 
ken_q_gather <- ken_q_gather %>% subset(���� != "���v")�@   # ���ނ����v�̂��̂��O���B
# ��s�̋斈��
ken_q_gather %>% subset(str_detect(�s�撬����,"��s.+��")) %>% ggplot() + 
  geom_col(aes(year_Q, �ː�, fill=����)) + facet_wrap(�s�撬����~.) +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

### ���ƁA�݉ƁA�����Z���csv�t�@�C�����쐬�E�ۑ� (�G�N�Z���p)
len <- length(ken_q)
mochi <- subset(ken_q, ����=="����")
kashi <- subset(ken_q, ����=="�݉�")
bunjo <- subset(ken_q, ����=="�����Z��")  
write.csv(mochi, "mochi.csv")
write.csv(kashi, "kashi.csv")
write.csv(bunjo, "bunjo.csv")