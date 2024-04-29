/** 
 * プログラム名：活動記録自動生成プログラム
 * プログラム説明：毎日２２時４０分ごろ、その日の活動記録を自動生成し、フォルダに格納してくれるプログラムです。
 * S先生（ID：jp2006）には対応していますが、他の指導者の方は未対応です。
 * 20人以上が入室した場合、21人目以降は存在しなかったものとして扱います。
 * 
 * 注意事項：活動計画書ができたら、URLを「活動計画書URL集」の中に記入しておいてください。
 * また、活動記録のテンプレート及び活動計画書エクセルではなく、スプレッドシートで保存しておいてください。
 * スプレッドシートの編集は可能ですが削除はしないでください。(2022/9/20)
 * 
 * BOXの制限人数を、20人から36人に変更。それに伴いコードとテンプレートを改良。(2022/2/20)
 * BOXの制限人数を撤廃(2022/10/5)
 * 
 * 作成者：ihiratch(junya737)
 * https://github.com/junya737
 * 最終更新日時：2022/2/20
*/


function Get_Document() {// 活動計画書からコマの時間情報term_timesを入手

  const today=new Date();//Dateオブジェクト取得
  const year=today.getFullYear();//年を取得
  const month=today.getMonth()+1;//月を取得
  const day =today.getDate();//日を取得
  const date=[year,month,day];//今日の日付情報　これ以降、説明の例として[2021,8,19]を使う

  const year_month=date[0]+"/"+date[1];//2021/8

  const url_collectionId="1kuovVpos0LMwgFuRqX009pecSqVhbI8TPLUAHsqpShE";//活動計画書URL集のID
  const sheetName="Sheet1";//シートの名称
  const url_collection_sheet=SpreadsheetApp.openById(url_collectionId).getSheetByName(sheetName);//URL集を開く

  const url_finder=url_collection_sheet.createTextFinder(year_month);//2021/8で検索
  const url_row=url_finder.findAll();//URLが記載されている行を取得
  const ssurl=url_collection_sheet.getRange(url_row[0].getRow(),2).getValue();//2021年8月の活動計画書のURL取得
  
  const search_word= date[1]+"月"+date[2]+"日";//検索ワード　8月19日
  const sheet=SpreadsheetApp.openByUrl(ssurl).getSheetByName(sheetName);//活動計画書取得

  const text_finder=sheet.createTextFinder(search_word);//8月19日で検索
  const cells=text_finder.findAll();//検索ワードが含まれるすべてのセルを取得
  

  let term_times=[];//コマの時間情報を入れる空配列　
  /*例
  [ [ 2021, 8, 19, 10, 0, 12, 0 ], 2021/8/19,10:00~12:00
  [ 2021, 8, 19, 15, 0, 17, 0 ],
  [ 2021, 8, 19, 18, 30, 20, 30 ] ]*/

  let activities=[];//活動内容 [ '自主練習', '合奏練習', '自主練習' ]


  let era_schedules=[];//元号とコマ割り
  
  //データの加工
  for(let i=0;i<cells.length;i++){

    if(sheet.getRange(cells[i].getRow(),3).getFontLine()=="line-through"){//取り消し線があればループに戻る。
      continue;
    }
    let stime=sheet.getRange(cells[i].getRow(),3).getDisplayValue();//コマ開始時間取得
    

    let etime=sheet.getRange(cells[i].getRow(),5).getDisplayValue();//コマ終了時間取得
    const activitiy = sheet.getRange(cells[i].getRow(),7).getDisplayValue();//
    
    stime=stime.split(":");// ：で分ける
    etime=etime.split(":");

    let stime_etime=stime.concat(etime).map( str => parseInt(str, 10) );//結合して数値化
    const term_time=date.concat(stime_etime);//コマの時間情報　[ 2021, 8, 19, 10, 0, 12, 0 ]

    const era ="令和"+(term_time[0]-2018)+"年";//元号
    let starting_time_doc=term_time[3]+":"+term_time[4]//コマ開始時間を文字列で取得
    let ending_time_doc=term_time[5]+":"+term_time[6];//コマ終了時間を文字列で取得

    //コマ割りが10:00~12:00のような時、10:0~12:0のようになるのを防ぐため、末尾に"０"を追加
    if(term_time[4]==0){
      starting_time_doc=starting_time_doc+"0";
      }

    if(term_time[6]==0){
      ending_time_doc=ending_time_doc+"0";
      }
    
    const schedule=term_time[1]+"月"+term_time[2]+"日"+starting_time_doc+"~"+ending_time_doc;//日時とコマ割り
    const era_schedule =[era,schedule];

    //空配列に追加
    term_times.push(term_time);
    activities.push(activitiy);
    era_schedules.push(era_schedule);

  }

  /*
  console.log(term_times);
  console.log(activities);
  console.log(era_schedules);*/

  return [activities,term_times,era_schedules];
  
}

function Copy_Seet_Auto(term_time) {//テンプレートをコピーしてフォルダ内に格納  　例：term_time=[ 2021, 8, 19, 10, 0, 12, 0 ]

  const mainfolderId="1KG-HwRqcSn4-f0nC9bOjb0y8vuLu0iPA";//活動記録フォルダID
  let month_name;//月の数

  if(term_time[1] <= 9){//月の数字を二桁にするための操作
    month_name = "0" + String(term_time[1]);
  } else {
      month_name= String(term_time[1]);
    }

  let starting_time_copy=String(term_time[3])+String(term_time[4])//コマ開始時間を文字列で取得
  let ending_time_copy=String(term_time[5])+String(term_time[6]);//コマ終了時間を文字列で取得

  //コマ割りが10:00~12:00のような時、100~120のようになるのを防ぐため、末尾に０を追加
  if(term_time[4]==0){
    starting_time_copy=starting_time_copy+"0";
    }

  if(term_time[6]==0){
    ending_time_copy=ending_time_copy+"0";
    }

  const folder_monthname=String(term_time[0]+"/"+month_name);//月のフォルダー名
  const folder_dayname = String(term_time[2]);//日のフォルダー名
  const temp_sheetId="1vCuvNzhpVe9NB9SWWbTXU7Hz4mB1cAOvbGvcTvxZL0Y";//テンプレートのファイルID
  const filename="【マンドリンオーケストラ・"+ term_time[1] +"月"+ term_time[2] +"日"+ starting_time_copy+"-"+ ending_time_copy +"】活動記録";//　ファイル名
  
  const temp_sheet= DriveApp.getFileById(temp_sheetId);//テンプレのファイルを取得
  const mainfolder=DriveApp.getFolderById(mainfolderId);//活動記録フォルダに移動
  const monthfolders=mainfolder.getFoldersByName(folder_monthname);//月の名前を持つフォルダを全てイテレータ形式で取得
  const monthfolder = monthfolders.next();//とりあえず１番目の要素を取得

  const dayfolders= monthfolder.getFoldersByName(folder_dayname);//日の名前を持つフォルダを全てイテレータ形式で取得
  const dayfolder = dayfolders.next();//とりあえず１番目の要素を取得

  const ss =temp_sheet.makeCopy(filename,dayfolder);//コピーを作成
  const ssId=ss.getId();//コピーのIDを取得

  return ssId;
}

function Automation() {//ID、体温等をシートに入力する。

  const [activities,term_times,era_schedules]=Get_Document();//activities：活動内容（自主練習など）
  //term_times: コマの時間情報[[ 2021, 8, 19, 18, 30, 20, 30 ] ,...]],  era_schedules:元号とコマ割り　[[令和３年,18:30~20:30],...]]取得

  let mail_texts=[];//メールの内容を入れる空配列
  const recipient="junya1702@gmail.com";//ihiratchのメアド
  const  subject="活動記録生成報告";//件名
  let body="ihiratch様\n\n";//本文

  if(term_times.length==0){//活動計画書にコマ情報がなかった場合
    const mail_text="本日は活動がありませんでした。";
    console.log(mail_text);
    body+=mail_text+"\n\n"+"活動記録自動生成コード";
    GmailApp.sendEmail(recipient,subject,body);//活動がなかった旨をメール送信

    return;//プログラム終了

  }
  const formId="1bc93Rbd6LsEAOYVdxzqeAHYe0A0b_6RHwSyFQwN-eLo";//入退室フォームのID
  const formName ="フォームの回答 1";
  const form = SpreadsheetApp.openById(formId).getSheetByName(formName);//フォームの取得
    
  const last_row=form.getLastRow();//フォーム回答の最終行を取得
  const rownum=2000; //取得するデータ数

  let timestamp=form.getRange(last_row-rownum,1,rownum).getDisplayValues()//フォーム回答のタイムスタンプ取得　形式上配列内配列 [["2021/08/09 14:06:16"],...]
  const original_data=form.getRange(last_row-rownum,2,rownum,5).getValues();//その他のデータ（名前、ID,入退室、体温）取得

  let Data=[];//加工後のデータを入れる空配列。
  const check="無し";//確認事項 デフォルトで「無し」に設定
  
  //あとあと使いやすいよう、Data= [[ 2021, 9, 2, 15, 48, 35, 948, 'ihiratch', 'G201', '入室', 36.6 ,"無し"],...]のような形式に加工
  for(let i=0;i<rownum;i++){
    
    timestamp[i]=timestamp[i][0].split(" ");// スペースで分割　["2021/08/09,14:06:16"]
    let date=timestamp[i][0].split("/");// 　"/"で分割[ '2021', '09', '02' ]
    let time=timestamp[i][1].split(":");//　　:で分割[ '15', '48', '35' ]
    let x_array=(date.concat(time).map( str => parseInt(str, 10) )); //　年月日と時刻を結合、文字列を数値化　[ 2021, 9, 2, 15, 48, 35 ]
    x_array.push(x_array[3]*60+x_array[4]);//　後で比較しやすくするため、分表示にした入室時間を挿入
    
    for(let j=0;j<4;j++){//体温等のデータも入れる
      x_array.push(original_data[i][j]);
    }
    x_array.push(check);//確認事項のデータも入れる。

    Data.push(x_array);//加工後のデータを入れ、ループに戻る
  }

  

  for(let k=0;k<term_times.length;k++){//コマの数だけループ

    const activity=activities[k];
    const term_time=term_times[k];
    const era_schedule=[era_schedules[k]];

    const ssId = Copy_Seet_Auto(term_time);// 活動記録スプレッドシートのIDを取得
    const sheetName ="活動記録";//シートの名称
    
    const sheet =SpreadsheetApp.openById(ssId).getSheetByName(sheetName);//スプレッドシート取得
    
    const start_time=term_time[3]*60+term_time[4];//比較しやすくするため、コマの開始時間を分表示　
    const end_time=term_time[5]*60+term_time[6];//コマの終了時間を分表示
    const width_time=15;//時間外での入室を許す時間幅　デフォルト15分
    const regulation_number=1000; //BOXの制限人数
    
    let satis_data=[];//条件を満たすデータを入れる空配列
    const shibata_ID="jp2006";
    let shibata_datas=[];// S先生用のデータ


    for(let i=0;i<rownum;i++){//Dataの中から、対象のコマに入室したデータのみ抽出
      //データが入室時のものであるか確認
      if(Data[i][9]=="入室"){
        //年月日が一致しているか確認
        if(Data[i][0]==term_time[0]&Data[i][1]==term_time[1]&Data[i][2]==term_time[2]){
          //入室時間がコマ内時間内にあるか確認
          if(Data[i][6]>=start_time-width_time&Data[i][6]<=end_time+width_time){
            //IDがS先生のものである場合
            if(Data[i][8]==shibata_ID){
              let shibata_data=[Data[i][7],Data[i][10],"℃","",Data[i][11]]
              shibata_datas.push(shibata_data);
              sheet.getRange(7,6,1,5).setValues(shibata_datas);//S先生のデータ入力

            }else{//S先生以外

            //データに重複がないかを確認
            
            let flag=1;//ブールフラッグ
            for(let j=0;j<satis_data.length;j++){
              //名前またはIDが一致している場合はflagを０にしてループから抜ける。
              if(satis_data[j][7]==Data[i][7]||satis_data[j][8]==Data[i][8]){
                flag=0;
                break;
              }
            }
            //フラッグが1ならデータを格納
            if(flag){
              satis_data.push(Data[i]);//条件を満たすデータを格納
              }
            
            

            

            //20人以上の入室が確認されたらループ終了
            if(satis_data.length>=regulation_number){
              break;
            }
            }
          }
        }
      }
    } 
    /**
     * console.log(satis_data);
    console.log(shibata_datas)
     */
    
    const text=era_schedule[0][0]+""+era_schedule[0][1]+"、"+activity+"のコマ";//ログ用

    if(satis_data.length==0){//そのコマが開かれていない場合ループに戻る
      const mail_text=text+"は解放されませんでした。"
      mail_texts.push(mail_text);//解放されなかった旨をメールに含める。
      console.log(mail_text);
  
      const file = SpreadsheetApp.openById(ssId)
      file.rename(file.getName()+"[活動なし]");
      continue;
    }

    let IDdatas=[];//IDを入れる空配列
    let otherdatas=[];//体温等を入れる空配列

    // 入力のためのデータ加工 
    for(let i=0;i<satis_data.length;i++){
      let IDdata=[satis_data[i][8]];//setValuesの仕様上、配列にしている。
      let otherdata=[satis_data[i][10],"℃","",satis_data[i][11]];

      IDdatas.push(IDdata);
      otherdatas.push(otherdata);
    }

    sheet.getRange(2,9,1,2).setValues(era_schedule);// 元号、コマ割りを入力
    sheet.getRange(11,2,satis_data.length,1).setValues(IDdatas);//IDの入力
    sheet.getRange(11,7,satis_data.length,4).setValues(otherdatas);//体温等の入力
    sheet.getRange(114,1).setValue(activity);//活動内容を入力

    const mail_text=text+"の活動記録を生成しました。";
    mail_texts.push(mail_text);
    console.log(mail_text);

    


  }

  
  for(let i=0;i<mail_texts.length;i++){//活動があった場合
    body += mail_texts[i]+"\n";
  }
  body+="\n活動記録自動生成コード"
  
  GmailApp.sendEmail(recipient,subject,body);//生成報告メール送信

}





