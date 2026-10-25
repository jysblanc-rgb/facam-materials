/* 패캠 강의 자료 공통: 탭, 브라우저 자동 저장(보조), 엑셀 파일 저장·불러오기 */
window.T = (function(){
  function load(key){ try{ return JSON.parse(localStorage.getItem(key)||"null"); }catch(e){ return null; } }
  function save(key,val){ try{ localStorage.setItem(key,JSON.stringify(val)); return true; }catch(e){ return false; } }
  function esc(s){ return String(s==null?"":s).replace(/[&<>"]/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c];}); }
  function toast(msg){
    var t=document.querySelector(".toast"); if(!t){ t=document.createElement("div"); t.className="toast"; t.setAttribute("role","status"); document.body.appendChild(t); }
    t.textContent=msg; t.classList.add("on"); clearTimeout(t._h); t._h=setTimeout(function(){t.classList.remove("on");},2200);
  }
  /* 탭: [data-tabs] 안의 button[data-tab] ↔ [data-panel] . 주소 뒤 #이름 으로 바로 열림 */
  function tabs(root,onShow){
    var btns=[].slice.call(root.querySelectorAll("[data-tab]"));
    function show(name,push){
      var hit=btns.some(function(b){return b.dataset.tab===name;}); if(!hit) name=btns[0].dataset.tab;
      btns.forEach(function(b){ var on=b.dataset.tab===name; b.setAttribute("aria-selected",on); var p=document.querySelector('[data-panel="'+b.dataset.tab+'"]'); if(p) p.hidden=!on; });
      if(push){ try{ history.replaceState(null,"","#"+name); }catch(e){} }
      if(onShow) onShow(name);
    }
    btns.forEach(function(b){ b.setAttribute("role","tab"); b.addEventListener("click",function(){ show(b.dataset.tab,true); }); });
    return show;
  }
  function xlsxReady(){ if(window.XLSX) return true; toast("엑셀 기능을 불러오는 중이에요. 잠시 뒤 다시 눌러 주세요."); return false; }
  function download(wb,name){ XLSX.writeFile(wb,name); }
  function pickFile(cb){
    var i=document.createElement("input"); i.type="file"; i.accept=".xlsx,.xls";
    i.onchange=function(){ var f=i.files[0]; if(!f) return; var r=new FileReader();
      r.onload=function(){ try{ cb(XLSX.read(new Uint8Array(r.result),{type:"array"})); }catch(e){ toast("이 파일은 열 수 없어요. 이 페이지에서 저장한 엑셀 파일을 골라 주세요."); } };
      r.readAsArrayBuffer(f); };
    i.click();
  }
  function rows(ws){ return XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:false,blankrows:true}); }
  return {load:load,save:save,esc:esc,toast:toast,tabs:tabs,xlsxReady:xlsxReady,download:download,pickFile:pickFile,rows:rows};
})();
