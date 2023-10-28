/**
 * Cookieのユーティリティクラス
 */
class CookieUtil {
  /**
   * 値を抽出
   * @param {string} cookie Cookieデータ（"name=value;...")
   * @return {string} value
   */
  static getValue(cookies, key) {
    const cookiesArray = cookies.split(';');

    for(const c of cookiesArray){
      const cArray = c.split('=');
      if(cArray[0] == key){
        return cArray[1]
      }
    }

    return false
  }
}
