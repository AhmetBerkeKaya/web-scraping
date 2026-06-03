const puppeteer = require('puppeteer');
const xlsx = require('xlsx');

async function startScraping() {
    console.log('1. Adım: TİMDER Tarayıcı başlatılıyor (Görsel Mod Açık)...');
    
    // Tarayıcıyı ekranda görmek ve olası engelleri aşmak için headless: false
    const browser = await puppeteer.launch({ 
        headless: false, 
        defaultViewport: null,
        args: ['--start-maximized'] 
    }); 
    
    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(90000); 

    const totalPages = 39; // Senin verdiğin 39 sayfa bilgisi
    const detailLinks = new Set(); 

    console.log('\n2. Adım: 39 Sayfadan "Detaylı Bilgi" linkleri toplanıyor...');

    // 1'den 39'a kadar tüm sayfaları dönüyoruz
    for (let i = 1; i <= totalPages; i++) {
        // TİMDER'in URL yapısına göre sayfayı oluştur
        const pageUrl = `https://www.timder.org.tr/uyeler/detayli-uye-arama/?sayfa=${i}&g=24`;

        console.log(`[Sayfa ${i}/${totalPages}] Taranıyor...`);
        
        try {
            await page.goto(pageUrl, { waitUntil: 'domcontentloaded' });
            await new Promise(r => setTimeout(r, 2000)); // Sayfanın yüklenmesi için 2 sn bekle

            // Sayfadaki "Detaylı Bilgi" butonlarının linklerini yakala
            const pageLinks = await page.evaluate(() => {
                const links = [];
                document.querySelectorAll('a').forEach(a => {
                    const text = a.innerText.toLowerCase();
                    const href = a.href;
                    // Eğer butonun üzerinde "detay" yazıyorsa veya linkin içinde "detay" kelimesi geçiyorsa al
                    if ((text.includes('detay') || href.includes('detay')) && href.includes('timder.org.tr')) {
                        // Kendi bulunduğu sayfayı link sanmasın diye ufak bir filtre
                        if (!href.includes('?sayfa=')) {
                            links.push(href);
                        }
                    }
                });
                return links;
            });

            // Bulunan linkleri ana torbaya at
            pageLinks.forEach(link => detailLinks.add(link));
            console.log(`-> Sayfa ${i} tamamlandı. Toplam toplanan potansiyel firma: ${detailLinks.size}`);
            
        } catch (error) {
            console.error(`-> ❌ Sayfa ${i} yüklenirken hata oluştu: ${error.message}`);
        }
    }

    const linksArray = Array.from(detailLinks);
    console.log(`\nTOPLAM ${linksArray.length} ADET FİRMA BULUNDU. Detaylar kazınıyor...\n`);

    if(linksArray.length === 0) {
         console.log("Hata: Üye linkleri toplanamadı. Lütfen sitenin yapısını kontrol et.");
         await browser.close();
         return;
    }

    const extractedData = [];

    // Detay sayfalarına girip veriyi sökme aşaması
    for (let i = 0; i < linksArray.length; i++) {
        const link = linksArray[i];
        console.log(`[${i + 1}/${linksArray.length}] Çekiliyor: ${link}`);
        
        try {
            await page.goto(link, { waitUntil: 'domcontentloaded' });
            await new Promise(r => setTimeout(r, 1000)); 
            
            const data = await page.evaluate(() => {
                // Firma Adını Başlıktan Al
                let firma = document.querySelector('h1') ? document.querySelector('h1').innerText.trim() : "";
                if (!firma || firma.length < 2) {
                    firma = document.querySelector('h2') ? document.querySelector('h2').innerText.trim() : document.title.split('-')[0].trim(); 
                }

                // Sayfadaki tüm yazıları satır satır böl
                const lines = document.body.innerText.split('\n').map(l => l.trim()).filter(l => l.length > 0);
                
                let tel = "-";
                let mail = "-";
                let web = "-";
                let adres = "-";
                let yetkili = "-";

                // Spagetti HTML Korumalı Akıllı Metin Okuyucu
                for (let i = 0; i < lines.length; i++) {
                    const line = lines[i];
                    const lowerLine = line.toLowerCase();
                    
                    // Yetkili Kişi Bulucu
                    if (lowerLine.startsWith('yetkili')) {
                        let val = line.replace(/^Yetkili[\s:.-]*/i, '').trim();
                        // Eğer isim "Yetkili" yazısının hemen yanında değil de bir alt satırındaysa
                        if (!val && lines[i+1]) val = lines[i+1].trim(); 
                        if (val && yetkili === "-") yetkili = val;
                    }

                    // Telefon Bulucu
                    if (lowerLine.startsWith('telefon') || lowerLine.startsWith('tel')) {
                        let val = line.replace(/^(Telefon|Tel)[\s:.-]*/i, '').trim();
                        if (!val && lines[i+1]) val = lines[i+1].trim();
                        if (val && val.length > 5 && val.length < 30 && tel === "-") tel = val;
                    }

                    // E-posta Bulucu
                    if (lowerLine.startsWith('e-posta') || lowerLine.startsWith('e-mail') || lowerLine.startsWith('email')) {
                        let val = line.replace(/^(E-posta|E-mail|Email)[\s:.-]*/i, '').trim();
                        if (!val && lines[i+1] && lines[i+1].includes('@')) val = lines[i+1].trim();
                        if (val && val.includes('@') && mail === "-") mail = val;
                    }

                    // Web Bulucu
                    if (lowerLine.startsWith('web') || lowerLine.startsWith('website')) {
                        let val = line.replace(/^(Web|Website)[\s:.-]*/i, '').trim();
                        if (!val && lines[i+1] && (lines[i+1].includes('www') || lines[i+1].includes('.com'))) val = lines[i+1].trim();
                        if (val && val.length > 4 && web === "-") web = val;
                    }

                    // Adres Bulucu
                    if (lowerLine.startsWith('adres')) {
                        let val = line.replace(/^Adres[\s:.-]*/i, '').trim();
                        if (!val && lines[i+1]) {
                            val = lines[i+1].trim();
                            // Adres uzunsa ve ikinci satıra taştıysa birleştir (Telefon/Mail satırına çarpmadığından emin olarak)
                            if (lines[i+2] && !lines[i+2].toLowerCase().includes('tel') && !lines[i+2].toLowerCase().includes('faks')) {
                                val += " " + lines[i+2].trim();
                            }
                        } else if (val) {
                             // Yanında yazıyorsa ve alt satıra da taştıysa
                             if (lines[i+1] && !lines[i+1].toLowerCase().includes('tel') && !lines[i+1].toLowerCase().includes('faks') && !lines[i+1].toLowerCase().includes('web')) {
                                 val += " " + lines[i+1].trim();
                             }
                        }
                        if (val && val.length > 5 && adres === "-") adres = val;
                    }
                }

                // Eğer e-posta metin okuyucudan kaçtıysa linklerin içinden zorla bul
                if (mail === "-") {
                    const mailLink = document.querySelector('a[href^="mailto:"]');
                    if (mailLink) mail = mailLink.href.replace('mailto:', '').trim();
                }

                // Aynı şekilde Web sitesi metin okuyucudan kaçtıysa dış linkleri tara
                if (web === "-") {
                    const allLinks = Array.from(document.querySelectorAll('a'));
                    for (let a of allLinks) {
                        const text = a.innerText.toLowerCase();
                        if (text.includes('www.') && !a.href.includes('timder.org.tr')) {
                            web = text;
                            break;
                        }
                    }
                }

                return {
                    "Firma Adı": firma || "-",
                    "Yetkili Kişi": yetkili,
                    "Telefon": tel,
                    "E-posta": mail,
                    "Web Sitesi": web,
                    "Adres": adres
                };
            });

            extractedData.push(data);

        } catch (err) {
            console.error(`Hata oluştu (${link}): Detay sayfası açılamadı.`);
        }
    }

    console.log('\n4. Adım: Excel dosyası oluşturuluyor...');
    
    if (extractedData.length > 0) {
        const worksheet = xlsx.utils.json_to_sheet(extractedData);
        const workbook = xlsx.utils.book_new();

        // 6 Sütun için Excel genişlik ayarları
        worksheet['!cols'] = [
            { wch: 45 }, // Firma Adı
            { wch: 25 }, // Yetkili Kişi
            { wch: 20 }, // Telefon
            { wch: 35 }, // E-posta
            { wch: 35 }, // Web Sitesi
            { wch: 70 }  // Adres
        ];

        xlsx.utils.book_append_sheet(workbook, worksheet, "Timder Full Veri");
        const fileName = `Timder_Uyeler_${Date.now()}.xlsx`;
        xlsx.writeFile(workbook, fileName);
        
        console.log(`✅ İşlem tamamlandı! Veriler "${fileName}" dosyasına kaydedildi.`);
    }

    await browser.close();
}

startScraping();