// Apps Script 번들에서는 현재 /exec 주소가 서버에서 주입됩니다.
window.TRAINING_SIGN_WEB_APP_URL = '';
window.TRAINING_SIGN_CONFIG = Object.freeze({
  API_URL: window.TRAINING_SIGN_WEB_APP_URL
});
