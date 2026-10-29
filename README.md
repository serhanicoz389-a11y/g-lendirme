# Dilek Ofis – QR Fiyat Kataloğu

Müşteri dükkândaki QR kodu telefonuyla okutur ve Dilek Ofis'in tüm kırtasiye ürünlerini
güncel fiyatlarıyla görür. Fiyatları bilgisayardaki yönetim panelinden istediğiniz zaman
değiştirirsiniz; değişiklik müşterinin ekranına anında yansır.

## Özellikler

**Müşteri sayfası** (QR ile açılır, telefona uygun)
- Kategorilere ayrılmış ürün ve fiyat listesi
- Ürün arama (ad veya barkod), kategori filtreleme
- **Barkod / karekod okutma:** "Okut" düğmesine basınca tam ekran kamera açılır; ürünün kodunu yeşil çerçevenin içine tutmak yeterli. Düz barkodlar (EAN-13, EAN-8, UPC, Code 128, Code 39…) ve kare kodlar (QR, DataMatrix, Aztec, PDF417, GS1) kendiliğinden okunur, ürün fiyatıyla açılır; listede olmayan ürün için "Ürün bulunamadı" yazar. Okuyucu ZXing'dir (`public/vendor/zxing-*`, MIT/Apache); bulanık görüntüde keskinleştirerek tekrar dener
- Satıştan kaldırdığınız ürünler görünmez
- **Sepet ve sipariş:** Müşteri ürünleri "Sepete ekle" ile (ya da barkodu okutup) sepete atar, siparişine bir ad verir
  (örn. "Okul listesi – Elif"), adını ve isteğe bağlı telefonunu yazıp onaylar. **Siparişlerim** bölümünde siparişin
  durumunu (Alındı, Hazırlanıyor, Hazır, Teslim edildi) görür. Arama bütün kelimelere bakar ve kaç çeşit bulunduğunu yazar.

**Yönetim paneli** (`/admin`, şifreli)
- **Ürünler & Fiyatlar:** Tabloda hücreye tıklayıp değiştirin, otomatik kaydedilir. Yeni ürün ekleme, silme, "Satışta" kutusuyla gizleme.
- **Barkod (isteğe bağlı):** Ürüne barkod yazılabilir veya USB barkod okuyucuyla okutulabilir. Arama kutusunda okutulan barkod kayıtlı değilse yeni ürün olarak eklemeyi önerir. Aynı barkod iki ürüne verilemez.
- **Kategoriler:** Yeni kategori ekleme, adını düzeltme, sırasını değiştirme, silme (içindeki ürünler başka kategoriye taşınır).
- **Toplu fiyat değiştirme:** Tüm ürünlere veya bir kategoriye % zam/indirim, istenen yuvarlamayla (50 kuruş, tam lira…).
- **Excel'e aktar / Excel'den yükle:** "Excel'e Aktar" gerçek bir Excel dosyası (.xlsx) indirir. Excel'de fiyatları,
  adları, kategorileri, barkodları değiştirin, yeni satır ekleyin ya da satırı silin, kaydedip "Excel'den Yükle" ile
  geri yükleyin. Uygulanmadan önce ne değişeceği gösterilir (değişen, yeni, silinecek ürünler ve sorunlu satırlar).
  Barkod sütunu metindir, Excel barkodu bozmaz. Son sütundaki "No"yu değiştirmeyin: ürünün adı değişse bile aynı ürün
  olarak tanınmasını sağlar. .csv dosyaları da okunur.
- **Karekod etiketi yazdır:** Listede görünen ürünler için A4'e etiket basar. Müşteri bu etiketi telefonunun kendi kamerasıyla okuttuğunda ürünün güncel fiyatı açılır (uygulama içinde "Okut" demeye gerek yok). Etikette fiyat yazmaz, böylece fiyat değişince yeniden basmak gerekmez.
- **Fiyat listesi yazdır:** Satıştaki tüm ürünlerin A4 fiyat listesi.
- **Siparişler:** Müşterilerin sitede onayladığı siparişler gelir (yeni siparişte ses ve sekmede sayı). Durumu tek
  tıkla değiştirilir, müşteri sitede görür; sipariş yazdırılır ya da Sipariş Formu'na aktarılır. İnternetteki siteden
  sipariş almak için bir kerelik kurulum: Ayarlar › İnternetten Sipariş Alma › Kurulum rehberi (gizli `siparisler` deposu
  ve yalnız Issues izni olan bir anahtar). Siparişler şifreli gider; bilgisayar kapalıyken gelenler depoda bekler.
- **Sipariş Formu:** Müşteri adı, ürünler, miktarlar, iskonto → numaralı, imza alanlı A4 sipariş formu çıktısı.
- **QR Kod:** Adres bir kez "sabitlenir" ve QR bir daha değişmez. Afiş olarak yazdırılır veya SVG indirilir. Bilgisayarın adresi değişirse panel ve siyah pencere uyarı verir.
- **Ayarlar:** Dükkân adı, telefon, adres (sayfada ve çıktılarda görünür).

## Kurulum (Windows)

1. ZIP'i bir klasöre çıkartın (ZIP'in içinden doğrudan çalıştırmayın).
2. `baslat.bat` dosyasına çift tıklayın. Yönetim paneli tarayıcıda açılır.
   - Bilgisayarda Node.js yoksa ilk açılışta otomatik olarak bu klasöre indirilir (~35 MB, kurulum gerekmez).
3. Varsayılan şifre: **dilek123** (aşağıdan değiştirin).

Siyah pencere açık kaldığı sürece sistem çalışır. Kapatırsanız müşteriler sayfayı açamaz.

### Şifreyi değiştirmek

`baslat.bat` dosyasını Not Defteri ile açıp `:calistir` satırının hemen altına ekleyin:

```
set ADMIN_PASSWORD=yeni-sifreniz
```

## Barkodla fiyat görme hakkında

Telefon tarayıcıları canlı kamerayı yalnızca güvenli (https) sitelerde açar.
Bu bir tarayıcı güvenlik kuralıdır, uygulamadan kapatılamaz. Uygulama dükkân Wi-Fi'ında
(`http://192.168…`) çalışırken müşteri "Barkod" düğmesine basınca **fotoğraf çekerek okutma** kullanılır:
telefonun kamerası açılır, barkodun fotoğrafı çekilir ve ürün fiyatıyla gösterilir.
Uygulama internette (https) yayınlandığında canlı kamera ile anında, otomatik okutma da çalışır.
Wi-Fi'da otomatik okutma için en pratik yol **karekod etiketleridir**: telefonun kendi kamerası karekodu
kendiliğinden tanır ve ürünün fiyat sayfasını açar.

## Müşteriler sayfaya nasıl ulaşır?

**Seçenek 1 – Dükkân Wi-Fi'ı (taslakta kullanılan yöntem):** Bilgisayar ve müşterinin telefonu
aynı Wi-Fi ağında olmalıdır. QR Kod sekmesi bilgisayarın yerel adresini (örn.
`http://192.168.1.20:3000/`) otomatik bulur. Modem ayarlarından bilgisayara sabit IP vermeniz önerilir;
yoksa adres değişebilir ve QR'ı yeniden basmanız gerekir. Windows ilk çalıştırmada güvenlik duvarı izni
sorarsa "İzin ver" deyin.

**Seçenek 2 – İnternette yayınlama (ücretsiz, önerilen):** Müşteri sayfası Cloudflare Pages üzerinde
`https://<ad>.pages.dev` adresinde yayınlanır. QR hiç değişmez, müşteri mobil veriyle de açar, barkodlar
canlı kamerayla okunur ve dükkân bilgisayarı kapalıyken de çalışır.

- Fiyatlar yine bu bilgisayardaki yönetim panelinden değiştirilir. Panel, değişiklikleri GitHub'daki
  `main` dalına gönderir; Cloudflare bunu görüp 1–2 dakika içinde siteyi günceller.
- "Değişiklikleri otomatik yayınla" açıksa son değişiklikten 15 saniye sonra kendiliğinden gönderilir;
  sağ üstteki **🌐 Yayınla** düğmesi durumu gösterir.
- Kurulum bir kereliğine ~10 dakikadır: yönetim paneli → **Ayarlar → İnternette Yayınlama → Kurulum rehberi**.
- İnternete yalnızca müşteri sayfası ve satıştaki ürünlerin listesi gönderilir; yönetim paneli ve şifre gönderilmez.
  GitHub anahtarı sadece bu bilgisayardaki `data/settings.json` dosyasında saklanır.

## Dosyalar

| Dosya | Görevi |
|---|---|
| `server.js` | Sunucu (ek paket gerektirmez) |
| `data/products.json` | Ürünler ve fiyatlar |
| `data/settings.json` | Dükkân bilgileri, sipariş form numarası (ilk kayıtta oluşur) |
| `public/index.html` | Müşteri sayfası |
| `public/admin.html` | Yönetim paneli |

**Yedek:** `data` klasörünü ara ara USB'ye veya buluta kopyalamanız yeterli.
Yönetim panelindeki "Excel'e Aktar" da yedek olarak kullanılabilir.

Taslaktaki ürün ve fiyatlar örnektir; kendi listenizi panelden veya Excel ile yükleyebilirsiniz.
