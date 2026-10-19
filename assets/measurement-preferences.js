(() => {
 const status=document.getElementById('measurementStatus');
 function show(){try{const signal=navigator.globalPrivacyControl===true || navigator.doNotTrack==='1';status.textContent=signal?'Website measurement is off because your browser sends a privacy signal.':localStorage.getItem('nolimitAnalyticsOptOut')==='1'?'Optional website measurement is off in this browser.':'Aggregate website measurement is allowed in this browser.';}catch{status.textContent='Optional website measurement is off because browser storage is unavailable.';}}
 document.getElementById('disableMeasurement').addEventListener('click',()=>{try{localStorage.setItem('nolimitAnalyticsOptOut','1');}catch{}show();});
 document.getElementById('enableMeasurement').addEventListener('click',()=>{try{localStorage.removeItem('nolimitAnalyticsOptOut');}catch{}show();});show();
})();
