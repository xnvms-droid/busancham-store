/**
 * 후원주점 서빙 시스템 — 접속 설정
 *
 * API_URL : 구글 시트 Apps Script를 [웹 앱]으로 배포한 뒤 발급된 주소를 따옴표 안에 붙여넣으십시오.
 *           예) 'https://script.google.com/macros/s/AKfycb.../exec'
 *           비워두면 첫 접속 화면에서 주소를 직접 입력할 수 있습니다(기기마다 입력 필요).
 * POLL_MS : 화면 자동 갱신 간격(밀리초). 기본 4000 = 4초. 포스 화면은 최대 3초 간격으로 갱신됩니다.
 */
window.APP_CONFIG = {
  API_URL: 'https://script.google.com/macros/s/AKfycbwGA9ysl0OcvhLSRs0Pf7p7G3BNFmyTsX_CAc4AzGPNEWGhqCLwQrSe_WZc4y04lIPG/exec',
  POLL_MS: 4000
};
