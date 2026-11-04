# Excel / Outlook sipariş yanıtlarına Gemini desteği

Gelen tedarikçi yanıtlarını Gemini ile yorumlar, doğrulanan sipariş kalemlerine yanıt, durum ve termin işler. Mail oluşturma modülü verilen dosyayla **bayt bayt aynıdır**. Mail gönderme, silme veya Gemini tarafından komut çalıştırma eklenmez. Kullanım: tabloyu doldur, `GelenYanitlariOku` makrosunu başlat.

## Windows'ta bir defalık kurulum

Gerekenler: masaüstü Excel, **klasik Outlook** ve açık Outlook profili, Python 3.10+ ve `py` launcher, internet ve Gemini API erişimi. Yeni Outlook, Excel Web ve macOS desteklenmez. Python harici paket istemez.

1. Sohbette paylaşılan API anahtarını iptal edip Google AI Studio'da yenisini oluşturun. Anahtarı sohbete veya depoya yazmayın.
2. Depoyu Windows bilgisayarına indirin/klonlayın. Excel'i tamamen kapatın. Bu klasörde PowerShell açıp `./Kurulum.ps1` çalıştırın. Kurumunuzun betik çalıştırma politikası engelliyorsa yöneticinizin onaylı yöntemini kullanın. Betik anahtarı gizli girişle sorar, sentetik siparişle **gerçek Gemini çağrısı** yapar ve başarılıysa kullanıcı ortam değişkenlerini ayarlar. E-posta verisi bu testte gönderilmez. Anahtar Windows kullanıcı ortamında saklanır; bu bilgisayardaki aynı kullanıcı süreçleri okuyabilir.
3. Mevcut Excel dosyanızın yedeğini alın ve `.xlsm` olarak kaydedin. `Alt+F11` ile VBA düzenleyicisini açın. Eski `OnlyReadReplies` modülünü dışa aktararak saklayın ve projeden kaldırın. `File > Import File` ile `macros/AcikSiparis_Yalniz_Gelen_Yanitlari.bas` dosyasını aktarın. Aynı isimde iki modül bırakmayın.
4. Mail oluşturma modülü zaten varsa **değiştirmeyin**. Yoksa `macros/Mail_Olusturma_Yalniz_Taslak.bas` dosyasını aktarın. Gelen dosyanın içeriği aynen korunmuştur.
5. `Debug > Compile VBAProject` çalıştırın. Okuma modülü Türkçe Windows VBA kodlamasıyla (Windows-1254), taslak modülü özgün kodlamasıyla saklanır. Türkçe olmayan Windows sisteminde metinleri ve derlemeyi ayrıca kontrol edin.
6. Excel'i yeniden açın. Mevcut `Ayarlar` sayfasındaki sipariş/tedarikçi sayfalarını ve sütunlarını kontrol edin. Önce **kopya çalışma kitabında** `GelenYanitlariOku` çalıştırın. İlk kurulumda eski makronun ayar soruları devam eder; sonraki çalışmalarda kayıtlı ayarlar kullanılır.

`GEMINI_BRIDGE_DIR` depo ana klasörüdür. Depoyu taşırsanız bu değişkeni güncelleyin. Varsayılan model `gemini-2.5-flash`; hesabınıza uygun model için `GEMINI_MODEL` değiştirilebilir. Özel Python yolu gerekiyorsa `GEMINI_PYTHON` içine yalnızca `python.exe` tam yolunu yazın. API anahtarını Excel hücresine, VBA'ya veya komut satırına koymayın.

## İşleyiş ve sınırlar

- Her çalıştırmadan önce kitap klasörüne `GeminiYedek_*.xlsm` (veya kitabın mevcut uzantısı) yedeği alınır. Klasör yazılabilir olmalıdır. Çalışma kitabı otomatik kaydedilmez; sonucu kontrol edip kaydedin.
- Yalnızca mevcut makronun seçtiği gelen kutusu, son 180 günlük iletiler ve tedarikçi listesindeki **tam e-posta eşleşmeleri** işlenir. Domain/ad benzerliği tek başına yazma izni vermez. Tedarikçinin alternatif adreslerini listeye ekleyin.
- Mevcut geçmiş-yanıt ayırıcısından çıkan yanıt metni, alınma tarihi ve yalnızca ilgili tedarikçinin aday satır/sipariş/kalem/malzeme kodları Google Gemini API'ye gönderilir. İş verilerinizin Google'a iletilmesini bu kullanım gerektirir. Ek dosyalar Gemini'ye gönderilmez; mevcut ek kaydetme davranışı korunur.
- Gemini sonucu JSON şemasına, aday satır listesine, yanıt içindeki gerçek alıntıya, benzersiz sipariş+kalem/malzeme eşleşmesine, tarih ve durum doğrulamasına tabidir. Model güven puanı tek başına doğruluk garantisi değildir. Kod eşleştirmeleri kasıtlı olarak sıkıdır; farklı sıfır dolguları gibi biçimler inceleme gerektirebilir.
- Belirsiz yanıtlar `Log` sayfasında `GEMINI_BELIRSIZ`, tam adresi bilinmeyen gönderenler `GEMINI_GONDEREN_BELIRSIZ` olarak bırakılır. Genel “hepsi haftaya” gibi yanıtlar bütün siparişlere yayılmaz. Tek sayfada en fazla 500 aday satır desteklenir.
- Bir satırın kayıtlı yanıt tarihi işlenen mailden yeni veya eşitse eski mail o satırı ezmez. Aynı anda çok sayıda mailde API çağrıları sıralı çalışır; Excel bekleyebilir ve API kullanım ücreti oluşabilir.
- API/JSON hatasında eski hatalı eşleştirmeye dönülmez; çalışma hata bildirerek durur. Önceki başarılı satırlar kitapta kalabilir; yedek geri dönüş içindir. Önceki `IslenenMailler` kayıtları korunur. Eski mailleri yeniden incelemek için önce kitap kopyasında ilgili kayıtları/yanıt tarihlerini değerlendirin; körlemesine temizlemeyin.
- Eski yanıt ayırma kuralı ve Gemini yanılabilir. Gerçek örneklerle doğrulama gereklidir. Kalıcı öğrenme veya tüm bilgisayarı yöneten bir ajan değildir; Excel/Outlook erişimini VBA gerçekleştirir, Gemini yalnızca metin yorumlar. Yanıtlar otomatik gönderilmez.

## Doğrulama

`python -m unittest discover -s tests -v`: aday/kalem ayrımı, belirsizlik, uydurma alıntı, geçersiz tarih/durum, tekrar eden satır, eksik anahtar, HTTP istek yapısı ve eksik model çıktısı için 15 test.

Bulut Linux ortamında bu testler geçti. Gerçek Gemini çağrısı ve Windows Excel/Outlook COM, VBA derlemesi ve gerçek çalışma kitabıyla uçtan uca test **bu ortamda yapılmadı**. Paylaşılan anahtar kullanılmadı veya depoya eklenmedi. `Kurulum.ps1` Windows'ta canlı Gemini testini çalıştırır; gerçek posta/Excel doğrulaması yukarıdaki kopya kitap adımında tamamlanmalıdır.
