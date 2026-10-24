# Dilek Ofis – QR Fiyat Kataloğu

Müşteri dükkândaki QR kodu telefonuyla okutur ve Dilek Ofis'in tüm kırtasiye ürünlerini
güncel fiyatlarıyla görür. Fiyatları bilgisayardaki yönetim panelinden istediğiniz zaman
değiştirirsiniz; değişiklik müşterinin ekranına anında yansır.

## Özellikler

**Müşteri sayfası** (QR ile açılır, telefona uygun)
- Kategorilere ayrılmış ürün ve fiyat listesi
- Ürün arama (ad veya barkod), kategori filtreleme
- **Barkod okutma:** "Barkod" düğmesine basıp ürünün barkodunu kameraya tutunca ürün ve fiyatı ekrana gelir
- Satıştan kaldırdığınız ürünler görünmez

**Yönetim paneli** (`/admin`, şifreli)
- **Ürünler & Fiyatlar:** Tabloda hücreye tıklayıp değiştirin, otomatik kaydedilir. Yeni ürün ekleme, silme, "Satışta" kutusuyla gizleme.
- **Barkod (isteğe bağlı):** Ürüne barkod yazılabilir veya USB barkod okuyucuyla okutulabilir. Arama kutusunda okutulan barkod kayıtlı değilse yeni ürün olarak eklemeyi önerir. Aynı barkod iki ürüne verilemez.
- **Kategoriler:** Yeni kategori ekleme, adını düzeltme, sırasını değiştirme, silme (içindeki ürünler başka kategoriye taşınır).
- **Toplu fiyat değiştirme:** Tüm ürünlere veya bir kategoriye % zam/indirim, istenen yuvarlamayla (50 kuruş, tam lira…).
- **Excel'e aktar / Excel'den yükle:** Listeyi indirip Excel'de düzenleyip geri yükleyebilirsiniz.
- **Fiyat listesi yazdır:** Satıştaki tüm ürünlerin A4 fiyat listesi.
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

Telefon tarayıcıları canlı kamerayı yalnızca güvenli (https) sitelerde açar. Uygulama dükkân Wi-Fi'ında
(`http://192.168…`) çalışırken müşteri "Barkod" düğmesine basınca **fotoğraf çekerek okutma** kullanılır:
telefonun kamerası açılır, barkodun fotoğrafı çekilir ve ürün fiyatıyla gösterilir.
Uygulama internette (https) yayınlandığında canlı kamera ile anında okutma da çalışır.

## Müşteriler sayfaya nasıl ulaşır?

**Seçenek 1 – Dükkân Wi-Fi'ı (taslakta kullanılan yöntem):** Bilgisayar ve müşterinin telefonu
aynı Wi-Fi ağında olmalıdır. QR Kod sekmesi bilgisayarın yerel adresini (örn.
`http://192.168.1.20:3000/`) otomatik bulur. Modem ayarlarından bilgisayara sabit IP vermeniz önerilir;
yoksa adres değişebilir ve QR'ı yeniden basmanız gerekir. Windows ilk çalıştırmada güvenlik duvarı izni
sorarsa "İzin ver" deyin.

**Seçenek 2 – İnternette yayınlama:** Müşteri mobil veriyle de açabilsin isterseniz uygulama
bir sunucuya (Render, Railway, VPS vb.) kurulur ve QR o adrese yönlendirilir. Bir sonraki adımda
birlikte karar verebiliriz.

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
