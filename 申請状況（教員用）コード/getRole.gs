function getRole(email) {
  if (isAuthorizedUser(email, ["事務員"])) return "事務員";
  if (isAuthorizedUser(email, ["教務主事"])) return "教務主事";
  // @ktc.ac.jpドメインを一般教員扱い
  if ((email || "").toLowerCase().endsWith("@ktc.ac.jp")) return "一般教員";
  // それ以外（外部）は未認可
  return "ゲスト";
}

/**
 * 指定したメールアドレスが roles のいずれかの権限を持っているか判定する
 */
function isAuthorizedUser(email, roles) {
  try {
    // --- (1) スクリプトプロパティから設定値を取得 ---
    const props       = PropertiesService.getScriptProperties();
    const sheetId     = props.getProperty('AUTH_SHEET_ID')     || '1ZL3LtXDaApR-6IbNJ3n1JtpeOob-Lv_RPWuk-pL6ZaA';
    const sheetName   = props.getProperty('AUTH_SHEET_NAME')   || '事務ー教務主事';
    const emailCol    = Number(props.getProperty('AUTH_EMAIL_COL'))    || 1;  // A列
    const roleCol     = Number(props.getProperty('AUTH_ROLE_COL'))     || 3;  // C列
    const cacheTtlSec = Number(props.getProperty('AUTH_CACHE_TTL'))     || 300;

    // --- (2) キャッシュから既存データを取得 ---
    const cache = CacheService.getScriptCache();
    let data = cache.get('AUTH_DATA');
    if (data) {
      data = JSON.parse(data);
    } else {
      const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
      if (!sheet) throw new Error(`シート「${sheetName}」が見つかりません`);
      data = sheet.getDataRange().getValues();
      cache.put('AUTH_DATA', JSON.stringify(data), cacheTtlSec);
    }

    // --- (3) 大文字小文字を区別しないよう正規化 ---
    const normalizedEmail = (email || '').toString().trim().toLowerCase();
    const normalizedRoles = roles.map(r => r.toLowerCase());

    // --- (4) データを先頭行ヘッダー除いて走査 ---
    for (let i = 1; i < data.length; i++) {
      const sheetMailRaw = data[i][emailCol - 1];
      const sheetMail    = (sheetMailRaw || '').toString().trim().toLowerCase();
      if (sheetMail === normalizedEmail) {
        const roleStr = (data[i][roleCol - 1] || '').toString();
        // ロール文字列をカンマで区切ってそれぞれ正規化
        const userRoles = roleStr
          .split(',')
          .map(r => r.trim().toLowerCase())
          .filter(r => r);

        // --- (5) 指定 roles と照合 ---
        for (const r of userRoles) {
          if (normalizedRoles.includes(r)) {
            Logger.log(`AUTHORIZED: ${email} as ${r}`);
            return true;
          }
        }
      }
    }

    Logger.log(`NOT AUTHORIZED: ${email}`);
    return false;

  } catch (e) {
    // --- (6) 何かエラーが起きても false を返しつつログを残す ---
    Logger.log(`Error in isAuthorizedUser: ${e.message}`);
    return false;
  }
}
