export const supabaseRequest=async(path:string,init:RequestInit={})=>{
 const url=process.env.SUPABASE_URL,key=process.env.SUPABASE_SECRET_KEY;
 if(!url||!key)throw new Error("SUPABASE_URL と SUPABASE_SECRET_KEY を設定してください");
 const response=await fetch(`${url}/rest/v1/${path}`,{...init,headers:{apikey:key,Authorization:`Bearer ${key}`,"Content-Type":"application/json",Prefer:"return=representation",...(init.headers||{})}});
 if(!response.ok)throw new Error(`Supabase error ${response.status}: ${await response.text()}`);
 return response.status===204?null:response.json();
};
