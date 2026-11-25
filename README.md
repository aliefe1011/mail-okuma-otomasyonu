# Açık Sipariş Mail Otomasyonu

Excel VBA ile açık siparişler için klasik Outlook'ta e-posta taslakları oluşturur ve tedarikçi yanıtlarını Excel'deki sipariş kalemleriyle eşleştirir.

| Modül | Görevi |
| --- | --- |
| [Mail_Olusturma_Yalniz_Taslak.bas](macros/Mail_Olusturma_Yalniz_Taslak.bas) | `OnlyCreateDrafts` modülü, siparişleri tedarikçiye göre gruplar ve Outlook Taslaklar'a kaydeder. Gönderim Outlook'tan elle yapılır. |
| [AcikSiparis_Tam_Gelismis_Filtre_v856.bas](macros/AcikSiparis_Tam_Gelismis_Filtre_v856.bas) | `OnlyReadReplies` modülü (8.5.6), gelen yanıtları eşleştirir ve ilgili sipariş satırlarına yazar. Önizleme, seçili e-posta ve gönderim takibi içerir. |

**Kendi verilerinizi kullanın.** Kaynaktaki `example.invalid` adresleri, örnek tedarikçi adları, sipariş ve malzeme numaraları yalnızca `MailOkumaOzTest` içindir. Bunları Excel listenize veya Outlook ayarlarınıza kopyalamayın. Gerçek adresleri, şirket içi alan adlarını ve çalışma dosyanızı kendi ortamınızda saklayın.

## 1. Makroları Excel'e ekleyin

1. Windows'ta masaüstü Excel ve **klasik Outlook** kullanın. Outlook'ta kullanacağınız posta hesabını açın.
2. Çalışma kitabınızın bir kopyasını alın ve **Excel Makro İçerebilen Çalışma Kitabı (.xlsm)** olarak kaydedin.
3. Yukarıdaki iki `.bas` dosyasını GitHub dosya sayfasından **Download raw file** ile indirin. Türkçe karakterler için dosyaların Windows-1254 kodlamasını koruyun.
4. Excel'de `Alt+F11` ile VBA düzenleyicisini açın. Aynı adlarda eski `OnlyReadReplies` veya `OnlyCreateDrafts` modülleri varsa önce dışa aktararak yedekleyin ve yerlerine yenilerini yüklemek için eski kopyaları çalışma kitabından kaldırın.
5. **File > Import File** ile iki dosyayı içe aktarın. Ardından **Debug > Compile VBAProject** ile projeyi derleyip kitabı kaydedin.
6. İsterseniz `Alt+F8` ile `MailOkumaOzTest` çalıştırın. Bu makro kurgusal kayıtları geçici bir kitapta kullanır ve Outlook'a bağlanmaz. Sonunda geçen kontrol sayısını veya takıldığı kontrolü gösterir.

## 2. Sipariş ve tedarikçi sayfalarını hazırlayın

1. Sipariş sayfalarınızda sütun başlıklarını ilk satıra yazın. Gerçek sipariş numarası, kalem numarası, tedarikçi adı ve malzeme kodunu ilgili sütunlara yerleştirin. Taslakta kullanılacak üretim yeri, açıklama, birim ve bakiye bilgilerini de kendi tablonuzda eşleyin. Birden çok sipariş sayfası kullanıyorsanız hepsinde aynı sütun düzenini kullanın.
2. Ayrı bir tedarikçi sayfasında tedarikçi unvanlarını ve güncel e-posta adreslerini girin. Sipariş tablosundaki adları tedarikçi listesiyle tutarlı yazın. Aynı tedarikçiye ait birden fazla adresi aynı hücrede noktalı virgülle ayırabilirsiniz.
3. Yanıt ve takip sonuçları için sipariş sayfalarında ayrı sütunlar ayırın. Böylece gelen yanıt, yanıt tarihi, teyitli termin ve diğer sonuçlar doğrudan kendi sonuç sütunlarına yazılır.

## 3. Ayarlar sayfasını doldurun

Mevcut `Ayarlar` sayfasını `Alt+F8 > AyarlariAc` ile görünür yapın. İlk kurulumda bu sayfa yoksa Excel'de `Ayarlar` adlı yeni bir çalışma sayfası oluşturun. Aşağıdaki hücreleri kendi kitabınıza göre doldurun.

| Hücre | Girilecek değer |
| --- | --- |
| `B1` | Sipariş sayfalarınızın tam adları. Birden fazla adı virgülle ayırın. |
| `B2` | Tedarikçi listenizin bulunduğu sayfanın tam adı. |
| `B3` | İlk veri satırının numarası. Başlık 1. satırdaysa genellikle `2`. |
| `B4` | Taranacak Outlook klasörü. Varsayılan Gelen Kutusu için boş bırakın. Başka klasör için kendi Outlook klasör yolunuzu girin. |
| `B5` | Hariç tutulacak şirket içi e-posta alan adları. Adresin `@` işaretinden sonraki kısmını yazın. Birden fazla alan adını noktalı virgülle ayırın. Boş bırakılırsa bu listeye göre şirket içi alan adı filtresi uygulanmaz. |
| `B7:B14` | Sırasıyla sipariş no, kalem no, tedarikçi, malzeme, üretim yeri, uzun metin, birim ve bakiye sütunlarının harfleri. |
| `B15:B20` | Sırasıyla mail atıldı, gelen yanıt, yanıt tarihi, yanıtlayan, eşleşme yöntemi ve ekler sütunlarının harfleri. |
| `B23:B24` | Tedarikçi sayfasında unvan ve e-posta sütunlarının harfleri. |
| `B26` | Sınırlı test çalıştırması için `EVET`, normal kullanım için `HAYIR`. |
| `B27:B28` | Tedarikçi yanıt durumu ve teyitli termin için ayrılan sütunların harfleri. |
| `B29` | Kaç gün geriye bakılacağı. Varsayılan `180`, kabul edilen aralık 1–3650. |
| `B30` | Gelen kutusunun alt klasörlerini de taramak için `EVET`, yalnızca seçili klasör için `HAYIR`. |
| `B31` | Eşleşen yanıtların eklerini diske kaydetmek için `EVET`, ek kaydetmeyi kapatmak için `HAYIR`. |

`Alt+F8 > AyarlariKaydet` ile kaydedin. Sayfa yeniden gizlenirse `AyarlariAc` ile açabilirsiniz. Makrolar bilinen başlıklara göre sütunları yeniden algılayabildiği için ekrandaki ayar özetini çalışma kitabınızla karşılaştırın.

## 4. Önizleyin ve çalıştırın

1. **Tedarikçi listesini kontrol edin.** `Alt+F8 > TedarikciAdresleriniOnizle` ile aktif siparişlerde sorunlu tedarikçi kayıtlarını bulun. Excel'deki `FiltreOnizleme` sayfasını kontrol edin.
2. **Taslak oluşturun.** `Alt+F8 > TaslaklariOlustur` komutunu çalıştırın. Outlook **Taslaklar** klasöründen alıcıyı, sipariş tablosunu ve metni kontrol edin. Sabit CC listesi boş gelir. CC gerekiyorsa kendi alıcılarınızı doğrudan Outlook taslağına ekleyin. Hazır olduğunda Outlook'tan gönderin.
3. **Yanıtları önizleyin.** `GelenYanitlariOnizle` ile taramayı başlatın veya Outlook'ta tek ileti seçip `GelenSeciliMailiOnizle` kullanın. Sipariş hücrelerine yazmadan bulunan adayları `FiltreOnizleme` sayfasına kaydeder.
4. **Yanıtları işleyin.** `GelenYanitlariOku`, ayarlı klasörleri tarayıp eşleşen satırları günceller. `GelenYanitlariDerinOku` alt klasörleri de tarar. Tek bir iletiyi yazmak için Outlook'ta o iletiyi seçin ve `GelenSeciliMailiIsle` çalıştırın.
5. **Sonuçları kontrol edin.** `Log` ve genel taramadaki `TerminRaporu` sayfalarını inceleyin. İşlenen e-postalar `IslenenMailler` sayfasında tutulur. Eşleşen ekler, ayar açıksa `YanitEkleri` klasörüne kaydedilir. Kontrolden sonra çalışma kitabını kaydedin.

## Yanıt eşleştirme

8.5.6, tedarikçiyle birlikte sipariş ve kalem çiftini doğrular. Aynı e-postadaki farklı siparişlere ait yanıtları ayırabilir. Uygun olduğunda malzeme kodundan, alıntılanan sipariş tablosundan veya Outlook gönderim zincirinden yararlanır. Belirsiz eşleşmeler günlükte incelenebilir. Kurallarla çıkarılan terminleri ve durumları çalışma kitabınıza kaydetmeden önce gözden geçirin.
