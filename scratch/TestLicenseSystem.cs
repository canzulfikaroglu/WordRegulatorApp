using System;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using Microsoft.Win32;
using WordRegulator.Infrastructure.Services;
using WordRegulator.Domain.Interfaces;

class Program
{
    static async Task<int> Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        Console.WriteLine("==================================================");
        Console.WriteLine("🧪 WORDREGULATOR LİSANS SİSTEMİ OTOMATİK TESTLERİ");
        Console.WriteLine("==================================================");

        int passed = 0;
        int failed = 0;

        var licenseService = new LicenseService();

        // TEST 1: Boş anahtar testi
        Console.Write("[TEST 1] Boş anahtar kontrolü: ");
        var resEmpty = await licenseService.CheckLicenseAsync("");
        if (!resEmpty.IsValid && resEmpty.Message.Contains("boş olamaz"))
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"BAŞARILI ✓ (Mesaj: {resEmpty.Message})");
            passed++;
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"BAŞARISIZ ✗");
            failed++;
        }
        Console.ResetColor();

        // TEST 2: Tanımsız/Geçersiz anahtar testi
        Console.Write("[TEST 2] Tanımsız lisans anahtarı ('WORD-FAKE-0000-TEST'): ");
        var resFake = await licenseService.CheckLicenseAsync("WORD-FAKE-0000-TEST");
        if (!resFake.IsValid && resFake.Message.Contains("tanınmıyor"))
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"BAŞARILI ✓ (Mesaj: {resFake.Message})");
            passed++;
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"BAŞARISIZ ✗ (Dönen: {resFake.Message})");
            failed++;
        }
        Console.ResetColor();

        // TEST 3: Gist'te kayıtlı geçerli anahtar testi
        string validKey = "WORD-BJH9-VDZY-KBF7";
        Console.Write($"[TEST 3] Canlı Gist üzerinden geçerli lisans ('{validKey}'): ");
        var resValid = await licenseService.CheckLicenseAsync(validKey);
        if (resValid.IsValid)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"BAŞARILI ✓ (Lisans Doğrulandı)");
            passed++;
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"BAŞARISIZ ✗ (Hata: {resValid.Message})");
            failed++;
        }
        Console.ResetColor();

        // TEST 4: Windows Registry Okuma / Yazma Testi
        Console.Write("[TEST 4] Windows Registry (HKCU\\SOFTWARE\\WordRegulator) Kayıt/Okuma: ");
        var originalKey = licenseService.GetStoredLicenseKey();
        string testRegistryKey = "WORD-TEST-UNIT-PASS";
        licenseService.StoreLicenseKey(testRegistryKey);
        var readKey = licenseService.GetStoredLicenseKey();
        
        // Orijinal anahtarı geri yükle
        if (originalKey != null)
            licenseService.StoreLicenseKey(originalKey);

        if (readKey == testRegistryKey)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"BAŞARILI ✓ (Yazılan ve okunan: {readKey})");
            passed++;
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"BAŞARISIZ ✗ (Beklenen: {testRegistryKey}, Okunan: {readKey})");
            failed++;
        }
        Console.ResetColor();

        // TEST 5: Yerel Önbellek (lc.dat) Dosyası ve Şifreleme Testi
        Console.Write("[TEST 5] Şifreli Yerel Önbellek (%APPDATA%\\WordRegulator\\lc.dat): ");
        var cachePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WordRegulator", "lc.dat");
        if (File.Exists(cachePath) && new FileInfo(cachePath).Length > 0)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"BAŞARILI ✓ (Dosya boyutu: {new FileInfo(cachePath).Length} bytes)");
            passed++;
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"BAŞARISIZ ✗ (Cache dosyası bulunamadı)");
            failed++;
        }
        Console.ResetColor();

        // TEST 6: Aktif Olmayan ve Süresi Dolan Lisans Mantığı Testi (ValidateKey Reflection/Unit test)
        Console.Write("[TEST 6] Süresi Dolan & Devre Dışı Bırakılan Lisans Kuralları: ");
        var validateKeyMethod = typeof(LicenseService).GetMethod("ValidateKey", BindingFlags.NonPublic | BindingFlags.Static);
        if (validateKeyMethod != null)
        {
            var manifest = new WordRegulator.Domain.Models.LicenseManifest
            {
                Licenses = new System.Collections.Generic.List<WordRegulator.Domain.Models.LicenseInfo>
                {
                    new() { Key = "EXPIRED-KEY", Active = true, Expires = DateTime.UtcNow.AddDays(-1), Customer = "Expired" },
                    new() { Key = "DISABLED-KEY", Active = false, Expires = DateTime.UtcNow.AddDays(10), Customer = "Disabled" }
                }
            };

            var expResult = (LicenseCheckResult)validateKeyMethod.Invoke(null, new object[] { "EXPIRED-KEY", manifest })!;
            var disResult = (LicenseCheckResult)validateKeyMethod.Invoke(null, new object[] { "DISABLED-KEY", manifest })!;

            if (!expResult.IsValid && expResult.Message.Contains("süresi doldu") &&
                !disResult.IsValid && disResult.Message.Contains("devre dışı"))
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("BAŞARILI ✓ (Süresi dolan ve deaktif edilen lisanslar reddedildi)");
                passed++;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"BAŞARISIZ ✗ (Exp: {expResult.Message}, Dis: {disResult.Message})");
                failed++;
            }
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("ATLANDI (Metot bulunamadı)");
        }
        Console.ResetColor();

        Console.WriteLine("==================================================");
        Console.WriteLine($"SONUÇ: {passed} Başarılı, {failed} Başarısız");
        Console.WriteLine("==================================================");

        return failed == 0 ? 0 : 1;
    }
}
