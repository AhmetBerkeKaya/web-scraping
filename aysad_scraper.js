const puppeteer = require('puppeteer');
const xlsx = require('xlsx');

async function startScraping() {
    console.log('1. Adım: AYSAD tarayıcı başlatılıyor (GÖRSEL MOD AÇIK)...');
    
    // headless: false yaparak tarayıcıyı görünür hale getirdik. Bot koruması varsa gözümüzle göreceğiz.
    const browser = await puppeteer.launch({ 
        headless: false, 
        defaultViewport: null, // Ekranı tam boyuta yayar
        args: ['--start-maximized'] 
    }); 
    
    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(90000); // Zaman aşımını 90 saniyeye çıkardık

    const detailLinks = new Set(); 
    const totalPages = 7; // Sayfa sayısını senin verdiğin bilgiye göre sabitledik

    console.log('\n2. Adım: Tüm sayfalardaki üye linkleri toplanıyor...');

    for (let i = 1; i <= totalPages; i++) {
        const currentUrl = i === 1 
            ? 'https://www.aysad.org.tr/uye-listesi/' 
            : `https://www.aysad.org.tr/uye-listesi/?page_3sdTA=${i}`;

        console.log(`Sayfa ${i} taranıyor: ${currentUrl}`);
        
        try {
            await page.goto(currentUrl, { waitUntil: 'domcontentloaded' });
            
            // Sayfanın JS ile kartları çizmesi için bekleme süresini 4 saniyeye çıkardık
            await new Promise(r => setTimeout(r, 4000)); 

            const pageLinks = await page.evaluate(() => {
                const links = [];
                document.querySelectorAll('a').forEach(a => {
                    const href = a.href;
                    // "/uye/" içeren ama "/uye-listesi/" (sayfalama) olmayan gerçek firma linklerini al
                    if (href && href.includes('/uye/') && !href.includes('uye-listesi')) {
                        links.push(href);
                    }
                });
                return links;
            });

            if (pageLinks.length === 0) {
                console.log(`-> ⚠️ DİKKAT: Sayfa ${i} boş geldi. Sayfa geç yüklenmiş veya site bot koruması açmış olabilir.`);
            } else {
                pageLinks.forEach(link => detailLinks.add(link));
                console.log(`-> Sayfa ${i} başarıyla tarandı. Toplam toplanan link: ${detailLinks.size}`);
            }
            
        } catch (error) {
            console.error(`-> ❌ Sayfa ${i} yüklenirken hata oluştu: ${error.message}`);
        }
    }

    const linksArray = Array.from(detailLinks);
    console.log(`\nTOPLAM ${linksArray.length} adet benzersiz üye profili bulundu.`);
    console.log('3. Adım: Profillerin detayları kazınıyor...\n');

    if(linksArray.length === 0) {
         console.log("Hata: Üye linkleri toplanamadı. Lütfen açılan tarayıcı penceresinde sitenin düzgün yüklenip yüklenmediğini kontrol et.");
         await browser.close();
         return;
    }

    const extractedData = [];

    for (let i = 0; i < linksArray.length; i++) {
        const link = linksArray[i];
        console.log(`[${i + 1}/${linksArray.length}] İşleniyor: ${link}`);
        
        try {
            await page.goto(link, { waitUntil: 'domcontentloaded' });
            await new Promise(r => setTimeout(r, 1000)); // Detay sayfasında da ufak bir nefes payı
            
            const data = await page.evaluate(() => {
                const extractValue = (searchText) => {
                    const allElements = Array.from(document.querySelectorAll('p, span, div, h1, h2, h3, strong, b'));
                    const labelNode = allElements.find(el => el.innerText && el.innerText.trim() === searchText);
                    
                    if (labelNode) {
                        if (labelNode.nextElementSibling) {
                            return labelNode.nextElementSibling.innerText.trim();
                        }
                        if (labelNode.parentElement && labelNode.parentElement.nextElementSibling) {
                            return labelNode.parentElement.nextElementSibling.innerText.trim();
                        }
                    }
                    return "-";
                };

                const h1 = document.querySelector('h1') ? document.querySelector('h1').innerText.trim() : "-";
                
                return {
                    "Firma Adı": extractValue('Firma') !== "-" ? extractValue('Firma') : h1,
                    "İsim": extractValue('İsim'),
                    "Telefon": extractValue('Telefon Numarası'),
                    "Mail": extractValue('E-mail Adres'),
                    "Web Sitesi": extractValue('Website URL'), 
                    "Adres": extractValue('Adres')           
                };
            });

            extractedData.push(data);

        } catch (err) {
            console.error(`Hata oluştu (${link}): Sayfa yüklenemedi.`);
        }
    }

    console.log('\n4. Adım: Excel dosyası oluşturuluyor...');
    
    const worksheet = xlsx.utils.json_to_sheet(extractedData);
    const workbook = xlsx.utils.book_new();

    worksheet['!cols'] = [
        { wch: 40 }, 
        { wch: 25 }, 
        { wch: 20 }, 
        { wch: 30 }, 
        { wch: 35 }, 
        { wch: 60 }  
    ];

    xlsx.utils.book_append_sheet(workbook, worksheet, "Aysad Full Veri");
    
    const fileName = `Aysad_Full_Uyeler_${Date.now()}.xlsx`;
    xlsx.writeFile(workbook, fileName);
    
    console.log(`✅ İşlem tamamlandı! Veriler "${fileName}" dosyasına kaydedildi.`);

    await browser.close(); // İşlem bitince tarayıcıyı kapatır
}

startScraping();