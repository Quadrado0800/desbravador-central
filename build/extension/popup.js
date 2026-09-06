const status=document.getElementById("status");
const button=document.getElementById("send");
function show(message){status.textContent=message}
button.addEventListener("click",async()=>{
 button.disabled=true; show("Coletando cookies...");
 try{
  const cookies=await chrome.cookies.getAll({domain:"desbravadorweb.com.br"});
  if(!cookies.length) throw new Error("Nenhum cookie encontrado para desbravadorweb.com.br.");
  const payload=cookies.map(c=>({
   name:c.name,value:c.value,domain:c.domain,path:c.path,
   secure:c.secure,httpOnly:c.httpOnly,expirationDate:c.expirationDate
  }));
  show(`Encontrados ${payload.length} cookies. Enviando...`);
  const response=await fetch("http://localhost:5000/api/session",{
   method:"POST",headers:{"Content-Type":"application/json"},
   body:JSON.stringify({cookies:payload})
  });
  const data=await response.json();
  if(!response.ok) throw new Error(data.error||`HTTP ${response.status}`);
  show(`Sessão conectada.\nCookies: ${data.cookie_count}`);
 }catch(err){show(`Erro: ${err.message}`)}
 finally{button.disabled=false}
});