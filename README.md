# Outlook gelen yanıtlarını Excel'e okuma — VBA 6.2

**Python, PowerShell, Gemini veya API anahtarı gerekmez.** Çalışan mail oluşturma modülüne dokunulmadı. Bu sürüm, yanıtların yanlış sipariş/kalem satırına yazılmasını hedefler.

## Bir defalık güncelleme

1. Mevcut Excel dosyanızın kopyasını alın; önce bu kopyada deneyin. Dosya `.xlsm` olmalı. Masaüstü Excel ve **klasik Outlook** gereklidir.
2. `macros/AcikSiparis_Yalniz_Gelen_Yanitlari.bas` dosyasını indirin. GitHub dosya sayfasında **Download raw file** düğmesini kullanın; web sayfasını `.bas` diye kaydetmeyin.
3. Excel'de `Alt+F11`. Soldaki proje ağacından **OnlyReadReplies** modülüne sağ tıklayın; eski modülü dışa aktararak yedekleyin ve kaldırın. Diğer modüllere, özellikle mail oluşturma modülüne dokunmayın.
4. `File > Import File` ile yeni `.bas` dosyasını ekleyin. `Debug > Compile VBAProject` çalıştırın. Türkçe VBA dışa aktarma kodlaması Windows-1254 korunur; Türkçe olmayan Windows'ta karakterleri ayrıca kontrol edin.
5. `Alt+F8` → **MailOkumaOzTest** → Çalıştır. Bu kontrol geçici bir kitap açar, örnek satırlarda eşleştirme ve gerçek hücre yazma kontrolleri yapar, kitabı kaydetmeden kapatır. Outlook'a bağlanmaz, gerçek sipariş verisini değiştirmez. Başarılıysa **23 kontrol gecti** mesajı beklenir.
6. Outlook açıkken `Alt+F8` → **GelenYanitlariOku** → Çalıştır. Kayıtlı `Ayarlar` sayfasını kontrol edip devam edin. Sonucu kopya kitapta doğruladıktan sonra asıl kitaba aynı modülü aktarın.

## Düzeltmeler

- Satırlar önce **aynı tedarikçiye** göre sınırlandırılır. Ad karşılaştırması harf, Türkçe karakter ve noktalama farklarını normalize eder; tedarikçi listesindeki ve sipariş tablosundaki unvanların bunun dışında tutarlı olması gerekir. Benzer ilk kelime tek başına yetmez.
- Sipariş numarasıyla ilk bulunan satıra yazılmaz. Aynı siparişin birden fazla kalemi varsa **kalem veya benzersiz malzeme kodu** gerekir. Belirsiz yanıt hiçbir satıra topluca yayılmaz; eski `UpdateAllLines / B6 = EVET` ayarı bu davranışı açmaz.
- Konudaki sipariş numarası, gövdedeki kalem bilgisiyle birlikte kullanılabilir. Tablo hücreleri ve HTML satır sonları korunur. `00010` ile `10` aynı kalem sayılır. Farklı kalemlerin yanıt metinleri ve tarihleri birbirine karıştırılmaz.
- Sipariş başlığı, ardından kalem ve termin ayrı satırlarda yazılmışsa önceki sipariş bağlamı korunur. Birden fazla kalemi aynı belirsiz cümlede belirten yanıtlar otomatik yayılmaz.
- Eski alıntı e-postalar satır başındaki gönderen başlığıyla ayrılır; yanıt içindeki sıradan `Konu:` / `Tarih:` kelimeleri metni kesmez. Agresif içerik filtresinin sipariş kodlarını/satır düzenini silmesi kaldırıldı.
- Tedarikçi hücresindeki noktalı virgül/virgülle ayrılmış bütün e-posta adresleri okunur. Çakışan adresler işlenmez. Gmail/Hotmail gibi ortak alan adları tedarikçi tanımak için kullanılmaz; gönderen-adı benzerliğiyle otomatik eşleştirme kaldırıldı.
- Geçersiz Outlook öğeleri kendi hata işleyicisiyle atlanır; bir önceki mailin bilgileri yeniden kullanılmaz. Klasör bulunamazsa yanlış klasöre sessizce devam edilmez.
- Eski tarihli yanıt, satırdaki yeni yanıtı ezmez. Çakışan sonuç/kaynak sütunu ayarları yazma başlamadan reddedilir.

## Kullanım ve sınırlar

- Günlük kullanım: tabloyu doldur, Outlook'u aç, **GelenYanitlariOku** makrosunu başlat. Mail oluşturma modülünüzü aynı şekilde kullanmaya devam edin.
- `Ayarlar` sayfasındaki sipariş sayfaları, tedarikçi sayfası, sütunlar ve gelen kutusu yolu doğru olmalı. Birden fazla sipariş sayfası aynı sütun düzenini kullanmalıdır. Tam yol gerekiyorsa `Posta Kutusu\Gelen Kutusu\Alt Klasör` biçimini kullanın.
- Yalnızca seçili gelen kutusu klasörü ve son 180 gün taranır; alt klasörler kendiliğinden taranmaz. Yeni Outlook ve Excel Web, VBA/Outlook COM akışını desteklemez.
- Belirsiz eşleşmeleri `Log` sayfasındaki **BELIRSIZ_YANIT / ESLESMEDI / BELIRSIZ_GONDEREN** kayıtlarında inceleyin. Belirsiz veya yanlış kaydedilmiş siparişleri tahmin ederek doldurmaz.
- Önceden işlenmiş mailler `IslenenMailler` sayfasında tutulur ve atlanır. Bu güncelleme geçmişte yanlış yazılmış satırları otomatik temizlemez. Yeniden deneme gerekiyorsa önce kitap kopyasında ilgili eski yanıt/tarihleri ve işlenmiş mail kayıtlarını inceleyin. `IslenenMailleriTemizle` bütün kayıtları temizler; asıl kitapta körlemesine çalıştırmayın.
- Makro sonuçları hücrelere yazar; otomatik mail göndermez. Mevcut ek kaydetme ve raporlama akışı korunmuştur. Sonuçları kontrol edip Excel'i kendiniz kaydedin.

## Kontrol durumu

Linux ortamında kaynak yapısı, prosedür adları/sınırları, VBA satır uzunlukları, değişen kodun kapsamı ve mail oluşturma dosyasının özgün dosyayla bayt eşitliği kontrol edildi. Windows için modüle **23 kontrol içeren `MailOkumaOzTest`** eklendi. **Bu kontroller burada çalıştırılmadı; Excel/VBA derlemesi ve gerçek Outlook uçtan uca testi de burada yapılamadı.** Çalışan Windows testi sonucu alınmadan makronun sorununuzu kesin çözdüğü iddia edilmez.
