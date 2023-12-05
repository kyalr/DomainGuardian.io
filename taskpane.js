/* eslint-disable prettier/prettier */
/* eslint-disable no-undef */
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      document.getElementById("app-body").style.display = "flex";
    }
  });