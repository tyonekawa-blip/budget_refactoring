// =========================================
// システム設定をスクリプトプロパティに移行するツール
// =========================================

// ※ MAIN_SS_ID は main.gs などで定義されている前提です
function migrateSettingsToProperties() {
  const ss = SpreadsheetApp.openById(MAIN_SS_ID);
  const sheet = ss.getSheetByName('システム設定');
  if (!sheet) {
    console.error("システム設定シートが見つかりません");
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const properties = PropertiesService.getScriptProperties();
  
  // 現在のプロパティを念のためクリア
  properties.deleteAllProperties();
  
  const config = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    const val = String(data[i][1]).trim();
    
    // キーが存在するものだけ保存
    if (key) {
      config[key] = val;
    }
  }
  
  // 新しく作成したJSON保存用フォルダのIDも追加でセットしておきます
  // （JSON用フォルダ：1zfAhdkywnh9ndN-JiTLG-5nsXPN4_IGa）
  config['JSON_CACHE_FOLDER_ID'] = '1zfAhdkywnh9ndN-JiTLG-5nsXPN4_IGa';
  
  // プロパティに一括保存
  properties.setProperties(config);
  
  console.log("✅ スクリプトプロパティへの移行が完了しました！");
  console.log(config);
}