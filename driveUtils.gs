// 改修版 driveUtils.gs
const BASE_FOLDER_ID = '1b6io1nnwUbjKNa_vz_yaQrRvm7a1ywIT';

function getOrCreateFolderByName(parentId, name) {
  return getOrCreateFolder(DriveApp.getFolderById(parentId), name);
}

function getOrCreateFolder(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function saveFilesToFolder(folder, base64Files, entryIndex, studentName) {
  const fileInfoList = [];

  // ─────────────────────
  // ① 入力の最終防御
  // ─────────────────────
  if (!Array.isArray(base64Files) || base64Files.length === 0) {
    Logger.log(`⚠️ [申請${entryIndex}] 添付ファイルなし、または配列ではありません`);
    return [];
  }

  for (let i = 0; i < base64Files.length; i++) {
    const f = base64Files[i];

    // オブジェクト構造チェック
    if (!f || typeof f !== 'object') {
      Logger.log(`⚠️ [申請${entryIndex}] 無効なファイルオブジェクト (index ${i})`);
      continue;
    }

    if (!f.data || typeof f.data !== 'string') {
      Logger.log(`⚠️ [申請${entryIndex}] base64 データ不正 (index ${i})`);
      continue;
    }

    Logger.log(`🔍 [申請${entryIndex}] file ${i + 1}: ${f.name || 'no-name'}, base64 length=${f.data.length}`);

    try {
      // ─────────────────────
      // ② base64 デコード
      // ─────────────────────
      const decoded = Utilities.base64Decode(f.data);

      if (!decoded || decoded.length === 0) {
        Logger.log(`⚠️ [申請${entryIndex}] デコード結果が空 (index ${i})`);
        continue;
      }

      // ─────────────────────
      // ③ ファイル名・拡張子
      // ─────────────────────
      const originalName = typeof f.name === 'string' ? f.name : 'file';
      const extension = getFileExtension(originalName) || '.bin';

      const renamed = `${studentName}_${entryIndex}_${i + 1}${extension}`;
      const safeName = sanitizeFilename(renamed);

      // MIME が壊れていても保存できるように保険
      const mimeType =
        typeof f.type === 'string' && f.type.trim() !== ''
          ? f.type
          : 'application/octet-stream';

      // ─────────────────────
      // ④ Drive 保存
      // ─────────────────────
      const blob = Utilities.newBlob(decoded, mimeType, safeName);
      const file = folder.createFile(blob);

      Logger.log(`✅ [申請${entryIndex}] 保存成功: ${file.getName()}`);

      fileInfoList.push({
        name: file.getName(),
        url: file.getUrl(),
        id: file.getId()
      });

    } catch (e) {
      // 1ファイル失敗しても他は継続
      Logger.log(`❌ [申請${entryIndex}] ファイル保存失敗 (${f.name || 'unknown'}): ${e.message}`);
    }
  }

  // ─────────────────────
  // ⑤ 最終チェック
  // ─────────────────────
  if (fileInfoList.length === 0) {
    Logger.log(`⚠️ [申請${entryIndex}] 有効な添付ファイルが1件も保存されませんでした`);
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

