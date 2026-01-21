# CSV to Excel Converter (CtE Converter)

Aplikasi desktop sederhana namun powerful untuk mengubah file CSV menjadi Excel (.xlsx) dengan cepat dan rapi.

![Screenshot Aplikasi](https://via.placeholder.com/600x400?text=CtE+Converter+Screenshot)

## Fitur Utama

- **Konversi Batch**: Proses banyak file CSV sekaligus.
- **Drag & Drop**: Cukup seret file CSV ke dalam aplikasi.
- **Auto-Formatting**: Kolom Excel otomatis dirapikan (Auto-width).
- **Kustomisasi Output**: Pilih simpan di folder sumber atau folder khusus.
- **Modern & Compact UI**: Tampilan bersih, ringan, dan mudah digunakan.
- **Bahasa Indonesia**: Antarmuka sepenuhnya dalam Bahasa Indonesia.

## Cara Instalasi / Penggunaan

1.  Download file executable terbaru (`CtE_Converter_v1_0_0_Final.exe`) dari menu Releases (jika ada) atau folder `dist`.
2.  Langsung jalankan file `.exe` tersebut (Portable, tidak perlu install).

## Cara Menjalankan dari Source Code

Jika Anda ingin mengembangkan ulang:

1.  Clone repository ini.
2.  Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```
3.  Jalankan aplikasi:
    ```bash
    python converter_app.py
    ```

## Cara Build (.exe)

Gunakan script `build_exe.bat` atau jalankan perintah:

```bash
python -m PyInstaller --noconsole --onefile --name "CtE_Converter_v1_0_0_Final" --collect-all ttkbootstrap --collect-all tkinterdnd2 converter_app.py
```

## Teknologi

- [Python](https://www.python.org/)
- [Tkinter](https://docs.python.org/3/library/tkinter.html) & [ttkbootstrap](https://ttbootstrap.readthedocs.io/) (UI)
- [Pandas](https://pandas.pydata.org/) (Data Processing)
- [OpenPyXL](https://openpyxl.readthedocs.io/) (Excel Formatting)
- [TkinterDnD2](https://github.com/pmgagne/tkinterdnd2) (Drag & Drop)

## Author

Dibuat oleh **Lalu Habib Sasiwimbe**.
CtE Converter v.1.0.0
