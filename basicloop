　Sub 反復動作()

      ‘*********************
      ‘Step1:入出力セルの設定
      ‘*********************
      ShName = ActiveSheet.Name
      Set CondSet = Sheets(ShName).Range("C2") ‘結果を計算するための条件を設定するためのセル
      Set Innput = Sheets(ShName).Range("B9") ‘引き込む条件が格納されているセル
      Set Result = Sheets(ShName).Range("C5") ‘結果が表示されるセル
      Set Output = Sheets(ShName).Range("D9") ‘結果を出力するためのセル
      Counter = 30 ‘反復数を設定

      For iii = 0 To Counter – 1
          ‘*********************
          ‘Step2:条件の引き込み
          ‘*********************
          aaa = Innput.Offset(iii, 0).Value ‘
          bbb = Innput.Offset(iii, 1).Value

          ‘*********************
          ‘Step3:条件の設定
          ‘*********************
          CondSet.Offset(0, 0).Value = aaa
          CondSet.Offset(1, 0).Value = bbb

          ‘*********************
          ‘Step4:計算
          ‘*********************
          Calculate

          ‘*********************
          ‘Step4:結果の出力
          ‘*********************
          Output.Offset(iii, 0).Value = Result.Offset(0, 0).Value

      Next iii

End Sub
