const puppeteer = require('puppeteer');
const xlsx = require('xlsx');

async function startScraping() {
    console.log('1. Adım: Tarayıcı başlatılıyor...');
    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(60000); 

    // BURAYI GÜNCELLE: Hangi derneği çekeceksen linkini yaz
    const URL = 'https://www.baca-der.org/uyelerimiz/'; 
    console.log(`2. Adım: ${URL} adresine gidiliyor...`);

    try {
        await page.goto(URL, { waitUntil: 'networkidle2' });
        await new Promise(resolve => setTimeout(resolve, 2000));

        const tableData = await page.evaluate(() => {
            const extracted = [];
            
            // BURAYI GÜNCELLE: Sitenin HTML'ine göre satırları (tr, div, kart vs) seç.
            const rows = Array.from(document.querySelectorAll('table tr')).slice(1);

            rows.forEach(row => {
                const cols = row.querySelectorAll('td');
                
                if (cols.length > 0) {
                    // Sütunlardaki verileri al, satır boşluklarını temizle
                    let rawFirma = cols[1] ? cols[1].innerText.trim() : "-";
                    let rawTel = cols[3] ? cols[3].innerText.trim().replace(/\n/g, ' ').replace(/\s+/g, ' ') : "-";
                    
                    // Baca-Der tablosundaki 5. index Adres, 6. index Web Sayfası
                    let rawAdres = cols[5] ? cols[5].innerText.trim().replace(/\n/g, ' ').replace(/\s+/g, ' ') : "-";
                    let rawWeb = cols[6] ? cols[6].innerText.trim() : "-";
                    
                    // =========================================================
                    // FİKS ÇIKTI ŞEMASI (6 SÜTUNLU)
                    // =========================================================
                    extracted.push({
                        "Firma Adı": rawFirma || "-",
                        "İsim": "-",         // Baca-Der'de isim sütunu yok
                        "Telefon": rawTel || "-",
                        "Mail": "-",         // Baca-Der'de mail sütunu yok
                        "Web Sitesi": rawWeb || "-",
                        "Adres": rawAdres || "-"
                    });
                }
            });

            return extracted;
        });

        console.log(`\nBaşarılı! Toplam ${tableData.length} firma verisi standart formata çekildi.`);
        console.log('3. Adım: Excel dosyası oluşturuluyor...');

        if (tableData.length > 0) {
            const worksheet = xlsx.utils.json_to_sheet(tableData);
            const workbook = xlsx.utils.book_new();
            
            // Excel sütun genişliklerini 6 sütuna göre ayarlıyoruz
            worksheet['!cols'] = [
                { wch: 40 }, // Firma Adı
                { wch: 25 }, // İsim
                { wch: 20 }, // Telefon
                { wch: 30 }, // Mail
                { wch: 35 }, // Web Sitesi
                { wch: 60 }  // Adres
            ];

            xlsx.utils.book_append_sheet(workbook, worksheet, "Dernek Üyeleri");
            
            const fileName = `Dernek_Uyeler_${Date.now()}.xlsx`;
            xlsx.writeFile(workbook, fileName);
            
            console.log(`✅ İşlem tamamlandı! Adres ve Web Sitesi dahil tertemiz 6 sütunlu veri "${fileName}" dosyasına kaydedildi.`);
        } else {
            console.log("❌ Hata: Veri bulunamadı. HTML seçicilerini (selectors) kontrol et.");
        }

    } catch (error) {
        console.error("Scraping sırasında hata:", error.message);
    } finally {
        await browser.close();
    }
}

startScraping();