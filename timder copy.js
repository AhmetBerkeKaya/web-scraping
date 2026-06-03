const puppeteer = require('puppeteer');
const xlsx = require('xlsx');

async function startTimderScrapingTest() {
    console.log('1. Adım: TIMDER tarayıcı başlatılıyor...');
    // İşlemi arka planda hızlandırmak istersen false değerini true yapabilirsin.
    const browser = await puppeteer.launch({ headless: false }); 
    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(0); // Uzun süreceği için timeout'u sınırsız yapıyoruz

    const totalPages = 39; // TAM SÜRÜM: 39 SAYFA
    const allDetailLinks = new Set(); 

    console.log(`\n2. Adım: ${totalPages} sayfalık firma linkleri toplanıyor...`);

    try {
        // SAYFA GEZME DÖNGÜSÜ (1'den 39'a kadar)
        for (let i = 1; i <= totalPages; i++) {
            const pageUrl = `https://www.timder.org.tr/uyeler/detayli-uye-arama/?sayfa=${i}&g=24`;
            console.log(`- Sayfa ${i} taranıyor...`);
            
            await page.goto(pageUrl, { waitUntil: 'networkidle2' });
            await new Promise(r => setTimeout(r, 1500)); 

            const pageLinks = await page.evaluate(() => {
                const links = [];
                document.querySelectorAll('a').forEach(a => {
                    if (a.href.match(/\/uye\/\d+\//)) {
                        links.push(a.href);
                    }
                });
                return links;
            });

            pageLinks.forEach(link => allDetailLinks.add(link));
        }

        const detailLinks = Array.from(allDetailLinks);
        console.log(`\nToplam ${totalPages} sayfada ${detailLinks.length} adet firma linki bulundu.\n`);
        
        const extractedData = [];

        console.log('3. Adım: Veriler DOM mantığı ile çekiliyor. Bu işlem biraz sürebilir, lütfen bekleyin...');
        for (let i = 0; i < detailLinks.length; i++) {
            const link = detailLinks[i];
            console.log(`[${i + 1}/${detailLinks.length}] Çekiliyor: ${link}`);
            
            try {
                await page.goto(link, { waitUntil: 'networkidle2' });
                await new Promise(r => setTimeout(r, 500)); 
                
                const data = await page.evaluate(() => {
                    let result = {
                        "Firma Adı": "-",
                        "Adres": "-",
                        "Yetkili Kişi": "-",
                        "Telefon": "-",
                        "E posta": "-",
                        "Website": "-"
                    };

                    const lines = document.body.innerText.split('\n').map(l => l.trim()).filter(l => l.length > 0);
                    for (let j = 0; j < lines.length; j++) {
                        if (lines[j].toLowerCase().startsWith("üye no:") && result["Firma Adı"] === "-") {
                            let val = lines[j].replace(/üye no[\s:]*\d+/i, '').trim();
                            if (val.length > 3) {
                                result["Firma Adı"] = val;
                            } else if (j + 1 < lines.length) {
                                result["Firma Adı"] = lines[j + 1].trim();
                            }
                            break;
                        }
                    }

                    const formRows = document.querySelectorAll('.row');
                    
                    formRows.forEach(row => {
                        const labelEl = row.querySelector('.formItem');
                        const inputContainer = row.querySelector('.formInput');

                        if (labelEl && inputContainer) {
                            const labelText = labelEl.innerText.toLowerCase().trim();
                            
                            const inputField = inputContainer.querySelector('input, textarea');
                            let value = "-";
                            
                            if (inputField) {
                                value = inputField.value || inputField.innerText || inputField.textContent;
                            } else {
                                value = inputContainer.innerText; 
                            }
                            
                            value = value.trim();
                            if(value === "") value = "-";

                            if (labelText.includes('adres')) {
                                result["Adres"] = value;
                            } else if (labelText.includes('yönetici')) {
                                result["Yetkili Kişi"] = value;
                            } else if (labelText.includes('telefon')) {
                                result["Telefon"] = value;
                            } else if (labelText.includes('e-mail') || labelText.includes('e-posta')) {
                                result["E posta"] = value;
                            } else if (labelText.includes('web') || labelText.includes('website')) {
                                result["Website"] = value;
                            }
                        }
                    });

                    return result;
                });
                
                extractedData.push(data);

            } catch (err) {
                console.error(`Hata oluştu (${link}): ${err.message}`);
            }
        }

        console.log('\n4. Adım: Tüm veriler için Excel dosyası oluşturuluyor...');
        
        if (extractedData.length > 0) {
            const worksheet = xlsx.utils.json_to_sheet(extractedData);
            const workbook = xlsx.utils.book_new();
            
            worksheet['!cols'] = [ 
                { wch: 50 }, // Firma Adı
                { wch: 60 }, // Adres
                { wch: 25 }, // Yetkili Kişi
                { wch: 25 }, // Telefon
                { wch: 35 }, // E posta
                { wch: 35 }  // Website
            ];

            xlsx.utils.book_append_sheet(workbook, worksheet, "Timder Tum Firmalar");
            const fileName = `Timder_Full_Veri_${Date.now()}.xlsx`;
            xlsx.writeFile(workbook, fileName);
            
            console.log(`✅ İşlem tamamlandı! Yaklaşık ${extractedData.length} firma çekildi. Excel dosyan: "${fileName}"`);
        } else {
            console.log("⚠️ Veri çekilemedi.");
        }

    } catch (error) {
        console.error("Genel bir hata oluştu:", error);
    } finally {
        await browser.close();
        console.log("Tarayıcı kapatıldı. Kod başarıyla sonlandı.");
    }
}

startTimderScrapingTest();