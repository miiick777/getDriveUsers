var SheetName = "getDriveUser";
activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName)

function getDriveUser() {

  firstStep();
  adminListAllTeamDrives();

}

function firstStep(sheetName) {

  //初期化する 
  activeSheet.clear();

  activeSheet.getRange(1, 1).setValue("ドライブ名")
  activeSheet.getRange(1, 1).setBackground("#7169e5");
  activeSheet.getRange(1, 1).setFontColor("#ffffff");

}  

function adminListAllTeamDrives(){
  //変数の宣言
  var pageTokenDrive;
  var pageTokenMember;
  var teamDrives;
  var permissions;

  //ドライブ名の一覧を取得
  do{
    teamDrives = Drive.Drives.list({pageToken:pageTokenDrive,maxResults:100,useDomainAdminAccess:true})
    if(teamDrives.items && teamDrives.items.length > 0){

      for (var i = 0; i < teamDrives.items.length; i++) {

        var teamDrive = teamDrives.items[i];
        //ドライブ名の一覧情報を転記
        activeSheet.getRange(i+2, 1).setValue(teamDrive.name)

          //ドライブごとのメンバーの権限を取得
          do{
            permissions = Drive.Permissions.list(teamDrive.id, {maxResults:40,pageToken:pageTokenMember,supportsAllDrives:true}) ;
            if(permissions.items && permissions.items.length > 0){
              for (var j = 0,k = 2; j < permissions.items.length; j++,k=k+2) {

              activeSheet.getRange(1, k).setValue("メンバー")
              activeSheet.getRange(1, k+1).setValue("権限")
              activeSheet.getRange(1, k, 1,k).setBackground("#7169e5");
              activeSheet.getRange(1, k, 1,k).setFontColor("#ffffff");


              //権限情報を取得して変数に格納
              var permission = permissions.items[j];
              activeSheet.getRange(i+2, k).setValue(permission.emailAddress)

              switch(permission.role){
              case "organizer":
                activeSheet.getRange(i+2, k+1).setValue("管理者")
                break;
              case "fileOrganizer":
                activeSheet.getRange(i+2, k+1).setValue("コンテンツ管理者")
                break;
              case "writer":
                activeSheet.getRange(i+2, k+1).setValue("投稿者")
                break;
              case "commenter":
                activeSheet.getRange(i+2, k+1).setValue("閲覧者(コメント可)")
                break;
              case "reader":
                activeSheet.getRange(i+2, k+1).setValue("閲覧者")
                break;
              }
            }
            }else{
              Logger.log("メンバー/権限が見つかりませんでした。");
            }

          //次のページのpageTokenを取得する
          pageTokenMember = permissions.nextPageTokens
          }while(pageTokenMember)
    }

    }else{
      Logger.log("共有ドライブが見つかりませんでした。");
    }

    //次のページのpageTokenを取得する
    pageTokenDrive = teamDrives.nextPageToken
    }while(pageTokenDrive)
}
