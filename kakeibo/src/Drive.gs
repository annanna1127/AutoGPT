/**
 * レシート画像を Googleドライブの「家計簿レシート/YYYY-MM」フォルダに自動仕分けして保存します。
 */

function getRootFolder_() {
  const id = props_().getProperty('DRIVE_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* 削除されていたら作り直す */ }
  }
  const it = DriveApp.getFoldersByName(CONFIG.DRIVE_ROOT_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.DRIVE_ROOT_NAME);
  props_().setProperty('DRIVE_FOLDER_ID', folder.getId());
  return folder;
}

function getMonthFolder_(ym) {
  const root = getRootFolder_();
  const it = root.getFoldersByName(ym);
  return it.hasNext() ? it.next() : root.createFolder(ym);
}

function safeName_(s) {
  return String(s || '').replace(/[\\/:*?"<>|\n\r]/g, '').slice(0, 40);
}

/** @return {{id: string, url: string}} */
function saveReceiptImage_(base64, mimeType, tx) {
  const date = tx ? formatDate_(tx.date) : today_();
  const ext = mimeType === 'image/png' ? 'png' : 'jpg';
  const name = [date, tx ? safeName_(tx.store) : '未読取', tx ? tx.amount + '円' : ''].filter(Boolean).join('_') + '.' + ext;
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, name);
  const file = getMonthFolder_(date.slice(0, 7)).createFile(blob);
  return { id: file.getId(), url: file.getUrl() };
}
