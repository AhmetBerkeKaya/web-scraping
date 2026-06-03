const puppeteer = require('puppeteer');
const xlsx = require('xlsx');

async function startScraping() {
    console.log('1. Adım: POMSAD tarayıcı başlatılıyor (Adres ve Web Sitesi dahil)...');
    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(60000); 

    const totalPages = 4;
    const detailLinks = new Set();

    console.log('\n2. Adım: Linkler toplanıyor...');

    for (let i = 1; i <= totalPages; i++) {
        const pageUrl = i === 1 
            ? 'https://pomsad.org.tr/uyelerimiz/' 
            : `https://pomsad.org.tr/uyelerimiz/page/${i}/`;

        console.log(`Sayfa ${i} taranıyor...`);

        try {
            await page.goto(pageUrl, { waitUntil: 'networkidle2' });
            await new Promise(r => setTimeout(r, 1000));

            const pageLinks = await page.evaluate(() => {
                const links = [];
                document.querySelectorAll('a').forEach(a => {
                    const href = a.href;
                    if (href.includes('/uyelerimiz/') && href.split('/uyelerimiz/')[1] && href.split('/uyelerimiz/')[1].length > 2) {
                        links.push(href);
                    }
                });
                return links;
            });

            pageLinks.forEach(link => detailLinks.add(link));
        } catch (err) {
            console.error(`Hata - Sayfa ${i}: ${err.message}`);
        }
    }

    const linksArray = Array.from(detailLinks);
    console.log(`\nToplam ${linksArray.length} firma bulundu. Detaylar çekiliyor...\n`);

    const extractedData = [];

    for (let i = 0; i < linksArray.length; i++) {
        const link = linksArray[i];
        console.log(`[${i + 1}/${linksArray.length}] Çekiliyor: ${link}`);
        
        try {
            await page.goto(link, { waitUntil: 'networkidle2' });
            
            const data = await page.evaluate(() => {
                // 1. Firma Adı
                let firma = document.querySelector('h1') ? document.querySelector('h1').innerText.trim() : "";
                if (!firma || firma.length < 2) {
                    firma = document.title.split('-')[0].trim(); 
                }

                const lines = document.body.innerText.split('\n').map(l => l.trim()).filter(l => l.length > 0);
                
                let tel = "-";
                let mail = "-";
                let web = "-";
                let adres = "-"; // ADRES DEĞİŞKENİ EKLENDİ

                // 2. Satır satır text okuma
                for (let line of lines) {
                    const lowerLine = line.toLowerCase();
                    
                    // Telefon bulucu
                    if (lowerLine.startsWith('telefon') || lowerLine.startsWith('tel')) {
                        let val = line.replace(/^(Telefon|Tel)[\s:.-]*/i, '').trim();
                        if (val && val.length > 5 && val.length < 25 && tel === "-") {
                            tel = val;
                        }
                    }

                    // Web bulucu
                    if (lowerLine.startsWith('web')) {
                        let val = line.replace(/^Web[\s:.-]*/i, '').trim();
                        if (val && val.length > 4 && web === "-") {
                            web = val;
                        }
                    }

                    // Adres bulucu (YENİ EKLENDİ)
                    if (lowerLine.startsWith('adres')) {
                        let val = line.replace(/^Adres[\s:.-]*/i, '').trim();
                        if (val && val.length > 5 && adres === "-") {
                            adres = val;
                        }
                    }
                }

                // 3. Mail ve Web için alternatif aramalar
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
                        if (text.includes('www.') && !href.includes('pomsad.org.tr')) {
                            web = text;
                            break;
                        } else if (href.startsWith('http') && !href.includes('pomsad.org.tr') && !href.includes('facebook') && !href.includes('twitter') && !href.includes('linkedin')) {
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
                    "Adres": adres // ADRES ÇIKTIYA EKLENDİ
                };
            });

            extractedData.push(data);

        } catch (err) {
            console.error(`Hata oluştu (${link})`);
        }
    }

    console.log('\n4. Adım: Excel dosyası oluşturuluyor...');
    
    if (extractedData.length > 0) {
        const worksheet = xlsx.utils.json_to_sheet(extractedData);
        const workbook = xlsx.utils.book_new();
        
        // Excel Sütun genişliklerini içeriğe göre ayarladık (Adres için geniş bir alan)
        worksheet['!cols'] = [ 
            { wch: 45 }, // Firma Adı
            { wch: 15 }, // İsim
            { wch: 25 }, // Telefon
            { wch: 35 }, // Mail
            { wch: 35 }, // Web Sitesi
            { wch: 60 }  // Adres (Uzun olabileceği için genişliği 60 yaptık)
        ];

        xlsx.utils.book_append_sheet(workbook, worksheet, "Pomsad Full Veri");
        const fileName = `Pomsad_Adresli_Veri_${Date.now()}.xlsx`;
        xlsx.writeFile(workbook, fileName);
        
        console.log(`✅ İşlem tamam! Adres dahil 6 sütunlu cillop gibi dosyan: "${fileName}"`);
    }

    await browser.close();
}

startScraping();