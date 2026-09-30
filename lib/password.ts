const encoder=new TextEncoder();
const hex=(bytes:ArrayBuffer|Uint8Array)=>Array.from(bytes instanceof Uint8Array?bytes:new Uint8Array(bytes)).map(x=>x.toString(16).padStart(2,"0")).join("");
const fromHex=(value:string)=>new Uint8Array(value.match(/.{1,2}/g)?.map(x=>parseInt(x,16))||[]);

export async function hashPin(pin:string,saltHex?:string){
 const salt=saltHex?fromHex(saltHex):crypto.getRandomValues(new Uint8Array(16));
 const key=await crypto.subtle.importKey("raw",encoder.encode(pin),"PBKDF2",false,["deriveBits"]);
 const bits=await crypto.subtle.deriveBits({name:"PBKDF2",salt,iterations:120000,hash:"SHA-256"},key,256);
 return{salt:hex(salt),hash:hex(bits)};
}
export async function verifyPin(pin:string,salt:string,expected:string){const result=await hashPin(pin,salt);let diff=result.hash.length^expected.length;for(let i=0;i<Math.min(result.hash.length,expected.length);i++)diff|=result.hash.charCodeAt(i)^expected.charCodeAt(i);return diff===0}
