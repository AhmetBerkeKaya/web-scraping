const puppeteer = require('puppeteer');
const xlsx = require('xlsx');

async function startScraping() {
    console.log('1. Adım: PUKAB Tarayıcı başlatılıyor (Görsel Mod Açık)...');
    
    const browser = await puppeteer.launch({ 
        headless: false, // Ne yaptığını görmek için açık bırakıyoruz
        defaultViewport: null,
        args: ['--start-maximized'] 
    }); 
    
    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(90000); 

    const URL = 'https://pukab.org.tr/uyelerimiz/';
    const detailLinks = new Set(); 

    console.log(`\n2. Adım: ${URL} adresinden üye linkleri toplanıyor...`);

    try {
        await page.goto(URL, { waitUntil: 'domcontentloaded' });
        await new Promise(r => setTimeout(r, 4000)); // Sayfanın yüklenmesi için bekle

        // PUKAB sayfasındaki üye detay linklerini yakala
        const pageLinks = await page.evaluate(() => {
            const links = [];
            // Sayfadaki tüm linkleri (a etiketlerini) dön
            document.querySelectorAll('a').forEach(a => {
                const href = a.href;
                // Genellikle logolara veya firma isimlerine verilen detay linklerini filtrele
                // PUKAB'da linkler /uyelerimiz/firma-adi/ veya /uye/firma-adi/ şeklinde olabilir
                if (href && (href.includes('pukab.org.tr') || href.includes('www.'))) {
                    // Sayfanın kendi linki (uyelerimiz) veya menü linkleri değilse listeye ekle
                    if (!href.endsWith('/uyelerimiz/') && !href.includes('iletisim') && !href.includes('hakkimizda')) {
                        links.push(href);
                    }
                }
            });
            return links;
        });

        // Bulunan tüm linkleri ana torbaya at (Set kullandığımız için mükerrer olanlar otomatik elenir)
        pageLinks.forEach(link => detailLinks.add(link));
        
        console.log(`-> Sayfa başarıyla tarandı. Toplam toplanan potansiyel link: ${detailLinks.size}`);
        
    } catch (error) {
        console.error(`-> ❌ Linkler toplanırken hata oluştu: ${error.message}`);
    }

    // Eğer linkler kendi sitesine (detay sayfasına) gitmiyorsa, logolar direkt firmaların kendi sitelerine (dış link) gidiyor demektir.
    // Dış link ise direkt "Web Sitesi" olarak alacağız.
    const linksArray = Array.from(detailLinks).filter(link => {
        // Eğer link pukab.org.tr içeriyorsa detay sayfasıdır, uzunluğunu kontrol et
        if (link.includes('pukab.org.tr')) {
            return link.split('/').filter(p => p.length > 0).length > 3; // Kök dizin dışındaki linkler
        }
        return true; // Dış linkleri de tut
    });

    console.log(`\nFiltreleme sonrası TOPLAM ${linksArray.length} adet geçerli üye profili bulundu.`);
    console.log('3. Adım: Profillerin detayları kazınıyor...\n');

    if(linksArray.length === 0) {
         console.log("Hata: Üye linkleri toplanamadı. Sitenin HTML yapısı çok farklı olabilir.");
         await browser.close();
         return;
    }

    const extractedData = [];

    for (let i = 0; i < linksArray.length; i++) {
        const link = linksArray[i];
        console.log(`[${i + 1}/${linksArray.length}] İşleniyor: ${link}`);
        
        try {
            // Eğer link dış bir siteyse (örneğin direkt adopen.com.tr'ye gidiyorsa)
            if (!link.includes('pukab.org.tr')) {
                 extractedData.push({
                    "Firma Adı": link.split('www.')[1] ? link.split('www.')[1].split('.')[0].toUpperCase() : link,
                    "İsim": "-",
                    "Telefon": "-",
                    "Mail": "-",
                    "Web Sitesi": link,
                    "Adres": "Sitede detay sayfası yok, doğrudan dış bağlantı."
                 });
                 continue; // Sonraki linke geç
            }

            await page.goto(link, { waitUntil: 'domcontentloaded' });
            await new Promise(r => setTimeout(r, 1500)); 
            
            const data = await page.evaluate(() => {
                let firma = document.querySelector('h1') ? document.querySelector('h1').innerText.trim() : "";
                if (!firma || firma.length < 2) {
                    firma = document.title.split('-')[0].trim(); 
                }

                const lines = document.body.innerText.split('\n').map(l => l.trim()).filter(l => l.length > 0);
                
                let tel = "-";
                let mail = "-";
                let web = "-";
                let adres = "-";

                // Metin Okuyucu (Spagetti HTML koruması)
                for (let i = 0; i < lines.length; i++) {
                    const line = lines[i];
                    const lowerLine = line.toLowerCase();
                    
                    if (lowerLine.startsWith('telefon') || lowerLine.startsWith('tel') || lowerLine.startsWith('t:')) {
                        let val = line.replace(/^(Telefon|Tel|T)[\s:.-]*/i, '').trim();
                        if (val && val.length > 5 && val.length < 25 && tel === "-") tel = val;
                    }

                    if (lowerLine.startsWith('web') || lowerLine.startsWith('w:')) {
                        let val = line.replace(/^(Web|W)[\s:.-]*/i, '').trim();
                        if (val && val.length > 4 && web === "-") web = val;
                    }

                    if (lowerLine.startsWith('adres') || lowerLine.startsWith('a:')) {
                        let val = line.replace(/^(Adres|A)[\s:.-]*/i, '').trim();
                        if (val && val.length > 5 && adres === "-") {
                            // Bazen adres iki satıra yayılır, o yüzden bir sonraki satırı da adrese ekleyelim
                            adres = val;
                            if (lines[i+1] && lines[i+1].length > 5 && !lines[i+1].toLowerCase().includes('tel') && !lines[i+1].includes('@')) {
                                adres += " " + lines[i+1];
                            }
                        }
                    }
                }

                const mailLink = document.querySelector('a[href^="mailto:"]');
                if (mailLink) {
                    mail = mailLink.href.replace('mailto:', '').trim();
                } else {
                    const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/;
                    const match = document.body.innerText.match(emailRegex);
                    if (match) mail = match[0];
                }

                if (web === "-") {
                    const allLinks = Array.from(document.querySelectorAll('a'));
                    for (let a of allLinks) {
                        const href = a.href;
                        const text = a.innerText.toLowerCase();
                        if (text.includes('www.') && !href.includes('pukab.org.tr')) {
                            web = text;
                            break;
                        } else if (href.startsWith('http') && !href.includes('pukab.org.tr') && !href.includes('facebook') && !href.includes('twitter') && !href.includes('instagram')) {
                             web = href;
                             break;
                        }
                    }
                }

                return {
                    "Firma Adı": firma || "-",
                    "İsim": "-",
                    "Telefon": tel,
                    "Mail": mail,
                    "Web Sitesi": web,
                    "Adres": adres
                };
            });

            extractedData.push(data);

        } catch (err) {
            console.error(`Hata oluştu (${link}): Sayfa yüklenemedi.`);
        }
    }

    console.log('\n4. Adım: Excel dosyası oluşturuluyor...');
    
    if (extractedData.length > 0) {
        const worksheet = xlsx.utils.json_to_sheet(extractedData);
        const workbook = xlsx.utils.book_new();

        worksheet['!cols'] = [
            { wch: 40 }, { wch: 20 }, { wch: 25 }, { wch: 35 }, { wch: 35 }, { wch: 70 }
        ];

        xlsx.utils.book_append_sheet(workbook, worksheet, "Pukab Full Veri");
        const fileName = `Pukab_Full_Uyeler_${Date.now()}.xlsx`;
        xlsx.writeFile(workbook, fileName);
        
        console.log(`✅ İşlem tamamlandı! Web Sitesi, Mail ve Adres dahil veriler "${fileName}" dosyasına kaydedildi.`);
    }

    await browser.close();
}

startScraping();