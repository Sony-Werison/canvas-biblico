const {chromium}=require('playwright');
const {readFile}=require('node:fs/promises');
const path=require('node:path');
const assert=require('node:assert/strict');
const JSZip=require('jszip');
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_PATH||'C:/Program Files/Google/Chrome/Application/chrome.exe'});
 try{
  const context=await browser.newContext({acceptDownloads:true});
  await context.route('http://local.test/**',async route=>{const name=new URL(route.request().url()).pathname.slice(1)||'index.html';const body=await readFile(path.join(__dirname,name));await route.fulfill({body,contentType:name.endsWith('.js')?'text/javascript':name.endsWith('.css')?'text/css':'text/html'});});
  const page=await context.newPage(), errors=[];page.on('pageerror',e=>errors.push(e.message));await page.goto('http://local.test/');
  await page.locator('#input').fill('No princípio era o Verbo');await page.locator('#add').click();assert.equal(await page.locator('.item').count(),5);
  await page.locator('.item').first().click();await page.locator('#apply-highlight').click();
  await page.locator('#note-text').fill('Uma anotação');await page.locator('#note').click();assert.equal(await page.locator('.item').count(),6);
  const [saved]=await Promise.all([page.waitForEvent('download'),page.locator('#save').click()]);const project=JSON.parse(await readFile(await saved.path(),'utf8'));assert.equal(project.items.length,6);assert.equal(project.items[0].highlight,'#ffe793');
  page.once('dialog',dialog=>dialog.accept());await page.locator('#file').setInputFiles({name:'projeto.canvasbiblico',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(project))});assert.equal(await page.locator('.item').count(),6);
  await page.locator('#original').click();assert.equal(await page.locator('.item').count(),5);await page.locator('#original').click();
  const [ppt]=await Promise.all([page.waitForEvent('download'),page.locator('#ppt').click()]);const zip=await JSZip.loadAsync(await readFile(await ppt.path()));assert.ok(zip.file('ppt/slides/slide1.xml'));assert.ok(zip.file('ppt/slides/slide2.xml'));assert.match(await zip.file('ppt/slides/slide2.xml').async('string'),/anotação/);
  await page.reload();assert.equal(await page.locator('.item').count(),6);assert.deepEqual(errors,[]);console.log('Importação, edição, salvamento, restauração e PowerPoint: OK');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
