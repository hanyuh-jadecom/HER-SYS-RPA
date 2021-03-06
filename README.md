# HER-SYS-RPA
新型コロナウイルス登録システム（HER-SYS）の登録を支援するプログラム

目的・動機
　院内でHER-SYSのデータ入力担当になりましたが、入力があまりにも面倒で驚きました。
　それまでは電子カルテから出力した一覧を保健所にFAXすれば済んでいたものが、WEB画面でちまちま入力しなければいけないことに違和感を覚えました。
　ちょうどSeleniumBasicで１本RPAを作ったばかりだったので、この作業もR自動化してみようと考えました。
　まだまだ不安定ですし、対応項目も少ないですが、「こんなこともできるんだ」と参考になると嬉しいと思い公開します。

制約条件
１．SeleniumBasicおよびChromeDrierが必要です。
２．カタカナ氏名と生年月日のみで、すでに登録されている患者かどうかの判断をしています。
　　同姓同名で生年月日が同じ患者は識別できません。
３．基本的に必須項目にしか対応していません。
４．患者の基本情報の登録と、検査内容、検査結果の登録のみを自動的に行います。
５．最低限のエラーチェックもできていない部分があり、不安定です。

使い方
１．Module1 の以下の部分を変更してください。
  Public Const User_ID As String = ""
  Public Const Password  As String = ""

  Public Const 外来機関 As String = "石岡第一病院"
  Public Const 保健所 As String = "土浦保健所"
  
２．BINフォルダのbook1.xsxの形式で、電子カルテなどからデータを抜き出してください。
　　CSI社のMIRAIs-PXからデータを抜き出すプログラムは作成済みですが、データベース構造の秘密保持契約の関係上、公開できません。
３．マクロ中のフォームを開き、２．で作ったファイルを指定して実行すると、ログインから自動的に動きます。

問い合わせ先
　羽生浩明（はにゅう　ひろあき）　hanyuh@jadecom.jp
メインの業務の間にプログラミングを行っているので、対応が遅くなることが多いかもしれませんが、ご容赦ください。
