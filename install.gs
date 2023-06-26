/**
 * 入力されたURLのフォルダにファイルを複製する関数
 */
function install() {
  const SRC_FILE_ID = PropertiesService.getScriptProperties().getProperty('SRC_FILE_ID');
  const spreadSheet = SpreadsheetApp.openById(SRC_FILE_ID);
  const fileId = spreadSheet.getId();
  const file = DriveApp.getFileById(fileId);
  const name = "集約さん-ファイル名を変更してください";

  let folderUrl = Browser.inputBox("集約さんを作りたいフォルダのURLを貼り付けてください。\n※マイドライブ直下は不可");

  const head1 = "https://drive.google.com/drive/folders/";
  const head2 = "https://drive.google.com/drive/u/0/folders/";
  const hip1 = "?usp=share_link";
  const hip2 = "?usp=sharing";

  if (folderUrl.includes(head1)){
    folderUrl = folderUrl.replace(head1,"");    
  }
  if (folderUrl.includes(head2)){
    folderUrl = folderUrl.replace(head2,"");    
  }
  if (folderUrl.includes(hip1)){
    folderUrl = folderUrl.replace(hip1,"");
  }
  if (folderUrl.includes(hip2)){
    folderUrl = folderUrl.replace(hip2,"");
  }
  const folderId = folderUrl;
  console.log(folderId);
  let folder;
  try{
    folder = DriveApp.getFolderById(folderId);
  } catch(e){
    Browser.msgBox("無効なURLです。URLと編集権限を確認してください。")
  }

  const userResponse = Browser.msgBox(`フォルダ：「${folder.getName()}」に集約さんを作ります。`,Browser.Buttons.OK_CANCEL);
  if(userResponse == "cancel"){
  Browser.msgBox("処理を終了します。");
  return;
  } else {
    try{
      file.makeCopy(name, folder);
    } catch(e){
      Browser.msgBox("エラー：URLと編集権限を確認してください。");
      return;
    }
  }
}