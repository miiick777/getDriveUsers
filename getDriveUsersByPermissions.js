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
  activeSheet.getRange(1, 2).setValue("管理者")
  activeSheet.getRange(1, 3).setValue("コンテンツ管理者")
  activeSheet.getRange(1, 4).setValue("投稿者")  
  activeSheet.getRange(1, 5).setValue("閲覧者(コメント可)")
  activeSheet.getRange(1, 6).setValue("閲覧者")
  activeSheet.getRange("A1:F1").setBackground("#7169e5");
  activeSheet.getRange("A1:F1").setFontColor("#ffffff");

}  

function adminListAllTeamDrives(){

  //変数の宣言
  var pageTokenDrive;
  var pageTokenMember;
  var teamDrives;
  var memberPermissions;
  //権限ごとのアドレスを格納
  var organizerData = '';
  var fileOrganizerData = '';
  var writerData = '';
  var commenterData = '';
  var readerData = '';

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
            memberPermissions = Drive.Permissions.list(teamDrive.id, {maxResults:40,pageToken:pageTokenMember,supportsAllDrives:true}) ;
            if(memberPermissions.items && memberPermissions.items.length > 0){
              for (var j = 0; j < memberPermissions.items.length; j++) {

              var permission = memberPermissions.items[j];
              //権限ごとに場合分けして変数に格納
              switch(permission.role){
              case "organizer":
                organizerData += permission.emailAddress + ',' + String.fromCharCode(10) ;
                break;
              case "fileOrganizer":
                fileOrganizerData += permission.emailAddress + ',' + String.fromCharCode(10) ;  
                break;
              case "writer":
                writerData += permission.emailAddress + ',' + String.fromCharCode(10) ;
                break;
              case "commenter":
                commenterData += permission.emailAddress + ',' + String.fromCharCode(10) ; 
                break;
              case "reader":
                readerData += permission.emailAddress + ',' + String.fromCharCode(10) ;  
                break;
              }

            }

            //権限ごとに格納した内容を転記  
            activeSheet.getRange(i+2,2).setValue(organizerData);
            organizerData = '';

            activeSheet.getRange(i+2,3).setValue(fileOrganizerData);
            fileOrganizerData = '';

            activeSheet.getRange(i+2,4).setValue(writerData);
            writerData = '';

            activeSheet.getRange(i+2,5).setValue(commenterData);
            commenterData = '';

            activeSheet.getRange(i+2,6).setValue(readerData);
            readerData = '';

            activeSheet.getRange(i+2,1,i+2,5).setWrap(true);
            activeSheet.getRange(i+2,1,i+2,5).setVerticalAlignment('top');
            activeSheet.getRange(i+2,1,i+2,5).setHorizontalAlignment('left');

            }else{
              Logger.log("メンバー/権限が見つかりませんでした。");
            }

          //次のメンバーのpageTokenを取得する
          pageTokenMember = memberPermissions.nextPageTokens

         }while(pageTokenMember)
    }

    }else{
     Logger.log("共有ドライブが見つかりませんでした。");
    }

    //次のドライブのpageTokenを取得する
    pageTokenDrive = teamDrives.nextPageToken
    }while(pageTokenDrive)

}
