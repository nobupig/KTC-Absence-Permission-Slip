// 改修版 driveUtils.gs
const BASE_FOLDER_ID = '1b6io1nnwUbjKNa_vz_yaQrRvm7a1ywIT';

function getOrCreateFolderByName(parentId, name) {
  return getOrCreateFolder(DriveApp.getFolderById(parentId), name);
}

function getOrCreateFolder(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function saveFilesToFolder(folder, base64Files, entryIndex, studentName){
  const fileInfoList = []; 

  if (!Array.isArray(base64Files) || base64Files.length === 0) {
    Logger.log("⚠️ 添付ファイルが空、または配列ではありません。");
    return [];
  }

  for (let i = 0; i < base64Files.length; i++) {
    const f = base64Files[i];
    Logger.log(`🔍 raw file info (index ${i}): ${JSON.stringify(f)}`);
    Logger.log(`   ↪ typeof f.data = ${typeof f.data}, length = ${f.data ? f.data.length : 'null'}`);

    if (!f || !f.data) {
      Logger.log(`⚠️ ファイルデータが無効なためスキップされました (index ${i})`);
      continue;
    }

    try {
      const decoded = Utilities.base64Decode(f.data);
      const extension = getFileExtension(f.name || 'file.pdf');
      const renamed = `${studentName}_${i + 1}${extension}`;
      const safeName = sanitizeFilename(renamed);

      const blob = Utilities.newBlob(decoded, f.type || 'application/octet-stream', safeName);
      const file = folder.createFile(blob);

      Logger.log(`✅ ファイル保存成功: ${file.getName()} (${file.getUrl()})`);
      fileInfoList.push({
        name: file.getName(),
        url: file.getUrl(),
        id: file.getId()
      }); 
    } catch (e) {
      Logger.log(`❌ ファイル保存失敗 (${f.name || `file_${i}`}): ${e.message}`);
    }
  }

  if (fileInfoList.length === 0) {
    Logger.log("⚠️ 保存されたファイルがありません。すべての処理に失敗した可能性があります。");
    return [];
  }

  return fileInfoList;
}

function getFileExtension(name) {
  const match = name.match(/\.[0-9a-zA-Z]+$/);
  return match ? match[0] : '';
}

function sanitizeFilename(name) {
  return name
    .replace(/[\\/:*?"<>|]/g, '')     // 禁止記号の除去（Windows準拠）
    .replace(/\s+/g, '_')             // 空白 → アンダースコア
    //.replace(/[^\x00-\x7F]/g, '')     // 非ASCII文字の除去（任意）
    .substring(0, 50);                // 長すぎる名前を50文字に切り詰め
}

