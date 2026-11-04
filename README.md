# Outlook gelen yanıtlarını Excel'e okuma — VBA 6.5

**Python, PowerShell, Gemini veya API anahtarı gerekmez.** Çalışan mail oluşturma modülüne dokunulmadı. Bu sürüm, yanıtların yanlış sipariş/kalem satırına yazılmasını hedefler.

## Bir defalık güncelleme

1. Mevcut Excel dosyanızın kopyasını alın; önce bu kopyada deneyin. Dosya `.xlsm` olmalı. Masaüstü Excel ve **klasik Outlook** gereklidir.
2. `macros/AcikSiparis_Yalniz_Gelen_Yanitlari.bas` dosyasını indirin. GitHub dosya sayfasında **Download raw file** düğmesini kullanın; web sayfasını `.bas` diye kaydetmeyin.
3. Excel'de `Alt+F11`. Soldaki proje ağacından **OnlyReadReplies** modülüne sağ tıklayın; eski modülü dışa aktararak yedekleyin ve kaldırın. Diğer modüllere, özellikle mail oluşturma modülüne dokunmayın.
4. `File > Import File` ile yeni `.bas` dosyasını ekleyin. `Debug > Compile VBAProject` çalıştırın. Türkçe VBA dışa aktarma kodlaması Windows-1254 korunur; Türkçe olmayan Windows'ta karakterleri ayrıca kontrol edin.
5. `Alt+F8` → **MailOkumaOzTest** → Çalıştır. Bu kontrol geçici bir kitap açar, örnek satırlarda eşleştirme ve gerçek hücre yazma kontrolleri yapar, kitabı kaydetmeden kapatır. Outlook'a bağlanmaz, gerçek sipariş verisini değiştirmez. Başarılıysa **67 kontrol gecti** mesajı beklenir.
6. Outlook açıkken `Alt+F8` → **GelenYanitlariOku** → Çalıştır. Kayıtlı `Ayarlar` sayfasını kontrol edip devam edin. Sonucu kopya kitapta doğruladıktan sonra asıl kitaba aynı modülü aktarın.

## Derin tarama (6.5)

Günlük kullanımda **`GelenYanitlariDerinOku`** makrosunu seçebilirsiniz. Seçili gelen kutusunu ve onun alt klasörlerini okur; başka posta kutularına veya çöp klasörlerine kendiliğinden geçmez. Alt klasörler dahil bütün öğeler birlikte alınma tarihine göre sıralanır. Normal `GelenYanitlariOku` yalnızca seçili klasörü okumaya devam eder. Her iki modda da son **180 gün** sınırı korunur. Büyük klasörlerde toplama/sıralama zaman ve bellek kullanır. Erişilemeyen klasör varsa eksik tarama başarılı sayılmaz; yazma başlamadan hata bildirilir. `DERIN_TARAMA` kaydı klasör ve öğe sayılarını gösterir.

Eşleştirme sırası:

1. Yeni yanıttaki açık sipariş/kalem veya benzersiz malzeme kodu.
2. Tablo hücreleri, `20 nolu kalem`, `Kalem (00020)` ve `sipariş/kalem` biçimleri.
3. Uzun metin sütunu ayarlıysa en az 20 karakterli, çok kelimeli **tam açıklama**; birkaç benzer kelime yeterli değildir. Aynı açıklama birden fazla satırda varsa otomatik seçim yapılmaz.
4. Yeni yanıt yetmezse en yakın **tek alıntılanmış mesajın** sipariş/kalem bağlamı. “Yarın teslim” gibi bir yanıt, alıntıda tek hedef varsa o hedefe bağlanır. Birden fazla kalem için “tüm kalemler” gibi açık kapsam ifadesi ve aynı sipariş şartı aranır. Konu, yeni yanıt ve alıntı arasında çelişki varsa tahmin edilmez.

Alıntı yalnızca **hangi satır** sorusuna yardımcı olur. Eski metindeki termin, durum ve açıklama yeni yanıt yerine yazılmaz. `From/Kimden/Gönderen`, “On … wrote” ve `>` alıntıları desteklenir. Farklı posta istemcilerinin bütün HTML düzenleri için garanti yoktur. Alıntı bağlamı en fazla 24.000 karakterle sınırlıdır. Başarı yöntemi hücrede `ALINTIDAN_SIPARIS_BAGLAMI` veya `ALINTIDA_TEK_HEDEF` olarak görülebilir.

Genel “Açık siparişler hakkında” başlıklarında Outlook `ConversationID` sağlıyorsa bağımsız konuşmalar ayrı tutulur; aynı konuşmanın en yeni yanıtı alınır. Açık sipariş numaralı aynı konu yine tedarikçi/konu bazında birleştirilir. ConversationID yoksa önceki konu gruplaması kullanılır.

PDF, resim/OCR, Excel veya diğer **ek dosyaların içeriği okunmaz**; mevcut ek kaydetme davranışı sürer. Mevcut işlenmiş mail kayıtları korunur. Bu sürümdeki Windows öz testi **67 kontrol** içerir; bu Linux ortamında çalıştırılmadı.

## Doğru sayaçlar ve eşleşmeme nedenleri

6.4 ve sonrasında “tedarikçisi tanınan yanıt”, “satır güncelleyen yanıt” ve “güncellenen satır” ayrı gösterilir. Belirsizlik ve tabloda tedarikçi bulunamaması mail başına sayılır; birden fazla sayfa aynı maili çoğaltmaz. Tedarikçisi bulunamayan mailler de atlanan toplamına dahildir.

`Log` sayfasında `SURUM=6.5`, `SUTUN_KONTROL` ve `ESLESME_NEDENI` kayıtları bulunur. Nedenler: `TEDARIKCI_SATIRI_YOK`, `KONU_SIPARISI_TEDARIKCIDE_YOK`, `BIRDEN_FAZLA_KALEM_ADAYI`, `YANITTA_SIPARIS_KALEM_MALZEME_YOK`, `KOD_VEYA_KALEM_ESLESMEDI`. Sayısal aday sayıları ve kullanılan sütunlar da yazılır. Daha yeni/eşit kayıt yüzünden yazılmayan satırlar `ESKI_VEYA_AYNI_TARIH`, özeti boş kalanlar `OZET_BOS` olarak ayrılır.

Bu değişiklik gözlenemeyen mail içerikleri için eşleşme tahmin etmez ve eşleştirme kurallarını gevşetmez. Loglar tek başına gerçek mail gövdesinin ve Excel satırının yerini tutmaz. Güncelleme geçmişte yanlış yazılmış hücreleri kendiliğinden onarmaz.

## Son yanıt, özet ve güncelleme

- Aynı **tedarikçi + konu** için mailler alınma tarihine göre yeniden eskiye sıralanır ve yalnızca en yeni normal yanıt seçilir. `RE:`, `FW:`, `YNT:` ve `[External]` gibi önekler temizlenir; farklı tedarikçi ve farklı sipariş numaralı konular birleştirilmez. Genel başlıklarda ConversationID mevcutsa bağımsız konuşmalar ayrılır; mevcut değilse konu gruplaması kullanılır. Boş başlıklar ileti kimliğiyle ayrılır. Eşit zaman damgalarında Outlook'un sıraladığı ilk ileti seçilir.
- En yeni ileti önceki çalışmada işlenmiş olsa da aynı grubun eski iletileri tekrar alınmaz. En yeni yanıt belirsizse eski maile dönerek hücre doldurulmaz. Otomatik yanıtlar ve teslim edilemedi bildirimleri başlık/ulaşılabilen mail üstbilgileriyle elenir; her sağlayıcının otomatik mesajını tanıma garantisi yoktur.
- Gelen yanıt hücresinde yalnızca **son yanıtın tarihli kısa özeti** görünür. Özet, gerçek metinden ilgili satırlar seçilerek çıkarılır; yapay zekâ veya API kullanılmaz. Selam, imza ve tekrarlanan satırlar ayıklanır; teslim/termin/durum/miktar ve olumsuz ifadeler korunmaya çalışılır. En fazla 600 karakterdir; uzun içerikte `...` görünür. Tam mail Outlook'ta kalır.
- `Mail Atıldı` sütununda `+` olması güncellemeyi engellemez. Daha yeni eşleşen yanıt geldiğinde özet, yanıt tarihi, gönderen, eşleşme yöntemi, durum, termin ve ekler güncellenir. `+` ve yanındaki mevcut yazı korunur. Yeni yanıtta kesin termin veya ek yoksa eski termin/ek yolu temizlenir; eski bilgi yeni yanıtın bilgisi gibi gösterilmez.
- Üzerine yazılan eski yanıt/durum/termin/ekler **aynı kitaptaki `YanitGecmisi`** sayfasına kaydedilir. Bu isimde farklı amaçlı bir sayfa varsa güncelleme durur; sayfayı koruyup yeniden adlandırın. Bir çalışma sırasında önceki satırlara yapılmış başarılı değişiklikler kalabilir; işlem bütün kitap için atomik değildir.
- `Pazartesi/Pazar` ve `Cumartesi/Cuma` karışması, gelecek haftaya fazladan yedi gün eklenmesi ve aynı günün gelecek yıla atlaması düzeltildi. Yalnızca “haftaya” veya ay adı verilmişse kesin gün uydurulmaz. Doğal dil yorumlaması kural tabanlıdır; karmaşık veya çelişkili tarih cümleleri insan kontrolü gerektirir.

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
- Normal modda seçili klasör; `GelenYanitlariDerinOku` modunda seçili klasör ve alt klasörler taranır. Son 180 gün sınırı korunur. Yeni Outlook ve Excel Web, VBA/Outlook COM akışını desteklemez.
- Belirsiz eşleşmeleri `Log` sayfasındaki **BELIRSIZ_YANIT / ESLESMEDI / BELIRSIZ_GONDEREN** kayıtlarında inceleyin. Belirsiz veya yanlış kaydedilmiş siparişleri tahmin ederek doldurmaz.
- Önceden işlenmiş mailler `IslenenMailler` sayfasında tutulur ve atlanır. Bu güncelleme geçmişte yanlış yazılmış satırları otomatik temizlemez. Yeniden deneme gerekiyorsa önce kitap kopyasında ilgili eski yanıt/tarihleri ve işlenmiş mail kayıtlarını inceleyin. `IslenenMailleriTemizle` bütün kayıtları temizler; asıl kitapta körlemesine çalıştırmayın.
- Makro sonuçları hücrelere yazar; otomatik mail göndermez. Mevcut ek kaydetme ve raporlama akışı korunmuştur. Sonuçları kontrol edip Excel'i kendiniz kaydedin.

## Kontrol durumu

Linux ortamında kaynak yapısı, prosedür adları/sınırları, VBA satır uzunlukları, değişen kodun kapsamı ve mail oluşturma dosyasının özgün dosyayla bayt eşitliği kontrol edildi. Windows için modüle **67 kontrol içeren `MailOkumaOzTest`** eklendi. **Bu kontroller burada çalıştırılmadı; Excel/VBA derlemesi ve gerçek Outlook uçtan uca testi de burada yapılamadı.** Çalışan Windows testi sonucu alınmadan makronun sorununuzu kesin çözdüğü iddia edilmez.
