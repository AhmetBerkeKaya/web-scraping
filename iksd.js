const puppeteer = require('puppeteer');
const xlsx = require('xlsx');

async function scrapeFullIksdMembers() {
    console.log('1. Adım: Tarayıcı başlatılıyor...');
    // İşlem uzun süreceği için arkada sessiz çalışsın istersen false'u true yapabilirsin
    const browser = await puppeteer.launch({ headless: false }); 
    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(0); // Dış siteler yavaş olabilir, süre sınırını kaldırdık

    console.log('\n2. Adım: İKSD Üyeler sayfasına gidiliyor ve firma linkleri toplanıyor...');
    await page.goto('https://www.iksd.com.tr/uyelerimiz/', { waitUntil: 'networkidle2' });

    // Sayfadaki dış linkleri (Firma web sitelerini) toplama
    const targetLinks = await page.evaluate(() => {
        const links = new Set();
        document.querySelectorAll('a').forEach(a => {
            const href = a.href;
            // http ile başlayan, kendi sitesi olmayan ve sosyal medya/gereksiz uzantılar içermeyen linkleri alıyoruz
            if (href && 
                href.startsWith('http') && 
                !href.includes('iksd.com.tr') && 
                !href.includes('facebook') && 
                !href.includes('instagram') && 
                !href.includes('twitter') && 
                !href.includes('linkedin') && 
                !href.includes('youtube') &&
                !href.includes('mailto:') && 
                !href.includes('tel:')) {
                links.add(href); 
            }
        });
        return Array.from(links);
    });

    console.log(`\nToplam ${targetLinks.length} adet farklı firma web sitesi bulundu. Dalgıç operasyonu başlıyor...\n`);

    const extractedData = [];

    // 3. Adım: Firmaların kendi sitelerini tek tek gezme
    for (let i = 0; i < targetLinks.length; i++) {
        const link = targetLinks[i];
        console.log(`[${i + 1}/${targetLinks.length}] Taranıyor: ${link}`);
        
        try {
            // Siteler çökük veya çok yavaş olabilir, 30 saniye içinde açılmazsa diğerine geçer
            await page.goto(link, { waitUntil: 'domcontentloaded', timeout: 30000 });
            await new Promise(r => setTimeout(r, 2000)); // Sayfadaki JS'lerin metinleri yüklemesi için kısa bekleme
            
            const data = await page.evaluate((currentLink) => {
                let result = {
                    "Firma Adı": "-",
                    "Telefon": "-",
                    "E posta": "-",
                    "Website": currentLink
                };

                // 1. Firma Adı: Sekme başlığından (title) temizleyerek alıyoruz
                let title = document.title || "";
                title = title.split('-')[0].split('|')[0].split(',')[0].trim();
                if (title) result["Firma Adı"] = title;
                else result["Firma Adı"] = currentLink.replace(/^https?:\/\/(www\.)?/, '').split('/')[0]; // Başlık yoksa domain adını al

                const bodyText = document.body.innerText;

                // 2. E-posta Avcısı (Regex)
                const emailRegex = /([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi;
                const emailMatch = bodyText.match(emailRegex);
                if (emailMatch) {
                    // Resim dosyalarını e-posta sanmasını engelliyoruz
                    const validEmails = emailMatch.filter(e => !e.toLowerCase().endsWith('.png') && !e.toLowerCase().endsWith('.jpg') && !e.toLowerCase().endsWith('.jpeg') && !e.toLowerCase().endsWith('.gif'));
                    if (validEmails.length > 0) result["E posta"] = validEmails[0];
                } else {
                    // Belki tıklanabilir mailto linki vardır
                    const mailto = document.querySelector('a[href^="mailto:"]');
                    if (mailto) result["E posta"] = mailto.href.replace('mailto:', '').trim();
                }

                // 3. Telefon Avcısı (Regex)
                const phoneRegex = /(?:\+90|0)?\s*[1-9]\d{2}\s*\d{3}\s*\d{2}\s*\d{2}|\b444\s*\d{1}\s*\d{3}\b|\b0\s*\(\s*[1-9]\d{2}\s*\)\s*\d{3}\s*\d{2}\s*\d{2}\b/g;
                const phoneMatch = bodyText.match(phoneRegex);
                if (phoneMatch) {
                    result["Telefon"] = phoneMatch[0].trim();
                } else {
                    // Belki tıklanabilir tel linki vardır
                    const telLink = document.querySelector('a[href^="tel:"]');
                    if (telLink) result["Telefon"] = telLink.innerText || telLink.href.replace('tel:', '').trim();
                }

                return result;
            }, link);

            extractedData.push(data);

        } catch (err) {
            console.error(`⚠️ Atlandı (${link}): Site çok yavaş, korumalı veya çökmüş olabilir.`);
            // Site hata verse bile boş olarak listeye ekleyelim ki linki kaybetmeyelim
            extractedData.push({
                "Firma Adı": "Erişim Hatası",
                "Telefon": "-",
                "E posta": "-",
                "Website": link
            });
        }
    }

    console.log('\n4. Adım: Tüm veriler toparlanıp Excel oluşturuluyor...');
    if (extractedData.length > 0) {
        const worksheet = xlsx.utils.json_to_sheet(extractedData);
        const workbook = xlsx.utils.book_new();
        
        worksheet['!cols'] = [ 
            { wch: 45 }, // Firma Adı
            { wch: 25 }, // Telefon
            { wch: 35 }, // E posta
            { wch: 40 }  // Website
        ];
        
        xlsx.utils.book_append_sheet(workbook, worksheet, "IKSD Firmalar");
        
        const fileName = `IKSD_Tum_Firmalar_${Date.now()}.xlsx`;
        xlsx.writeFile(workbook, fileName);
        console.log(`✅ İşlem Jilet Gibi Tamamlandı! Dosyan: "${fileName}"`);
    } else {
        console.log("⚠️ Veri çekilemedi.");
    }

    await browser.close();
    console.log("Tarayıcı kapatıldı. Geçmiş olsun!");
}

scrapeFullIksdMembers();