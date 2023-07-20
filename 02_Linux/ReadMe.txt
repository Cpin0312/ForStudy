==================================================
■初回環境準備
==================================================
●Java8 JREインストール
--------------------------------------------------
<\\hsadfs33.hsad.hitachi-solutions.co.jp\cru22\プロジェクト3\西武SPC\tool\java\linux>
jre-8u202-linux-x64.rpm

--------------------------------------------------
install 手順
--------------------------------------------------

１．jre8がインストールされていないことを確認。
ll /usr/java

※jre1.8.0_202-amd64が表示されないこと。


２．パッケージをインストールする。(オプション：-i install, -v information, -h installprocess)
sudo rpm -ivh jre-8u202-linux-x64.rpm


３．Javaパスの確認
ll /usr/java

※jre1.8.0_202-amd64が表示されること。

==================================================
●JSVCインストール
--------------------------------------------------
<\\hsadfs33.hsad.hitachi-solutions.co.jp\cru22\プロジェクト3\西武SPC\tool\java\linux>
apache-commons-daemon-jsvc-1.0.13-7.el7.x86_64.rpm

１．JSVCをインストールする

sudo rpm -ivh apache-commons-daemon-jsvc-1.0.13-7.el7.x86_64.rpm

２．インストール確認

ll /usr/bin/jsvc

※/usr/bin/jsvcが表示されること

==================================================
