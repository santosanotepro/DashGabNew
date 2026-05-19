import json, base64, math, pandas as pd

_ASSETS_LOADED = False
try:
    import os, sys
    _dir = os.path.dirname(os.path.abspath(__file__))
    if _dir not in sys.path:
        sys.path.insert(0, _dir)
    from _peta_assets import GEO_B64, KREF_B64, KWP_B64, JS_B64
    _ASSETS_LOADED = True
except ImportError:
    GEO_B64 = KREF_B64 = KWP_B64 = JS_B64 = ""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
"""
Generator LHA DJP - Gabungan
============================
Satu input file Excel LHA ->
Satu HTML output dengan tab [Dashboard Monitoring] dan [Peta Sebaran LHA]

Requirements: pip install pandas openpyxl
Jalankan: python Generator_LHA_DJP_Gabungan.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json, base64, os, sys, threading, traceback

# =============================================================
# PETA ASSETS (dari _peta_assets.py di folder yang sama)
# =============================================================
_ASSETS_LOADED = False
try:
    _dir = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, _dir)
    from _peta_assets import GEO_B64, KREF_B64, KWP_B64, JS_B64
    _ASSETS_LOADED = True
except ImportError:
    GEO_B64 = KREF_B64 = KWP_B64 = JS_B64 = ""

# =============================================================
# KOLOM WAJIB DASHBOARD
# =============================================================
DASH_REQUIRED_COLS = {
    'PENERBIT':         'Penerbit LHA (KPDJP/KANWIL)',
    'TAHUN_PAJAK':      'Tahun Pajak',
    'POTENSI_LHA':      'Nilai Potensi LHA',
    'KPP':              'Kode KPP (3 digit)',
    'NAMA_KPP':         'Nama KPP',
    'KWL':              'Kode Kanwil',
    'KANWIL':           'Nama Kanwil',
    'Status SP2DK':     'Status SP2DK (Terbit/Belum Terbit SP2DK)',
    'POTENSI_AWAL':     'Nilai Potensi Awal',
    'Status LHP2DK':    'Status LHP2DK (Terbit/Belum Terbit LHP2DK)',
    'POTENSI_AKHIR':    'Nilai Potensi Akhir',
    'NILAI_PEMBAYARAN': 'Nilai Pembayaran/Realisasi',
    'KATEGORI':         'Kategori LHA',
    'JENIS_LHA':        'Jenis LHA',
}


# =============================================================
# KPP MAPPING
# =============================================================
KPP_MAPPING = [
    {'kwl':10,'kwl_name':'Kanwil DJP Aceh','kpp':101,'kpp_name':'KPP Pratama Banda Aceh'},
    {'kwl':10,'kwl_name':'Kanwil DJP Aceh','kpp':104,'kpp_name':'KPP Pratama Bireuen'},
    {'kwl':10,'kwl_name':'Kanwil DJP Aceh','kpp':105,'kpp_name':'KPP Pratama Langsa'},
    {'kwl':10,'kwl_name':'Kanwil DJP Aceh','kpp':102,'kpp_name':'KPP Pratama Lhokseumawe'},
    {'kwl':10,'kwl_name':'Kanwil DJP Aceh','kpp':103,'kpp_name':'KPP Pratama Meulaboh'},
    {'kwl':10,'kwl_name':'Kanwil DJP Aceh','kpp':106,'kpp_name':'KPP Pratama Tapak Tuan'},
    {'kwl':10,'kwl_name':'Kanwil DJP Aceh','kpp':107,'kpp_name':'KPP Pratama Subulussalam'},
    {'kwl':10,'kwl_name':'Kanwil DJP Aceh','kpp':108,'kpp_name':'KPP Pratama Aceh Besar'},
    {'kwl':20,'kwl_name':'Kanwil DJP Sumatera Utara I','kpp':123,'kpp_name':'KPP Madya Medan'},
    {'kwl':20,'kwl_name':'Kanwil DJP Sumatera Utara I','kpp':119,'kpp_name':'KPP Pratama Binjai'},
    {'kwl':20,'kwl_name':'Kanwil DJP Sumatera Utara I','kpp':111,'kpp_name':'KPP Pratama Medan Barat'},
    {'kwl':20,'kwl_name':'Kanwil DJP Sumatera Utara I','kpp':112,'kpp_name':'KPP Pratama Medan Belawan'},
    {'kwl':20,'kwl_name':'Kanwil DJP Sumatera Utara I','kpp':122,'kpp_name':'KPP Madya Dua Medan'},
    {'kwl':20,'kwl_name':'Kanwil DJP Sumatera Utara I','kpp':124,'kpp_name':'KPP Pratama Medan Petisah'},
    {'kwl':20,'kwl_name':'Kanwil DJP Sumatera Utara I','kpp':121,'kpp_name':'KPP Pratama Medan Polonia'},
    {'kwl':20,'kwl_name':'Kanwil DJP Sumatera Utara I','kpp':113,'kpp_name':'KPP Pratama Medan Timur'},
    {'kwl':20,'kwl_name':'Kanwil DJP Sumatera Utara I','kpp':125,'kpp_name':'KPP Pratama Lubuk Pakam'},
    {'kwl':30,'kwl_name':'Kanwil DJP Sumatera Utara II','kpp':126,'kpp_name':'KPP Pratama Sibolga'},
    {'kwl':30,'kwl_name':'Kanwil DJP Sumatera Utara II','kpp':127,'kpp_name':'KPP Pratama Balige'},
    {'kwl':30,'kwl_name':'Kanwil DJP Sumatera Utara II','kpp':114,'kpp_name':'KPP Pratama Tebing Tinggi'},
    {'kwl':30,'kwl_name':'Kanwil DJP Sumatera Utara II','kpp':128,'kpp_name':'KPP Pratama Kabanjahe'},
    {'kwl':30,'kwl_name':'Kanwil DJP Sumatera Utara II','kpp':115,'kpp_name':'KPP Pratama Kisaran'},
    {'kwl':30,'kwl_name':'Kanwil DJP Sumatera Utara II','kpp':118,'kpp_name':'KPP Pratama Padang Sidempuan'},
    {'kwl':30,'kwl_name':'Kanwil DJP Sumatera Utara II','kpp':117,'kpp_name':'KPP Pratama Pematang Siantar'},
    {'kwl':30,'kwl_name':'Kanwil DJP Sumatera Utara II','kpp':116,'kpp_name':'KPP Pratama Rantau Prapat'},
    {'kwl':40,'kwl_name':'Kanwil DJP Riau','kpp':218,'kpp_name':'KPP Madya Pekanbaru'},
    {'kwl':40,'kwl_name':'Kanwil DJP Riau','kpp':211,'kpp_name':'KPP Pratama Pekanbaru Senapelan'},
    {'kwl':40,'kwl_name':'Kanwil DJP Riau','kpp':216,'kpp_name':'KPP Pratama Pekanbaru Tampan'},
    {'kwl':40,'kwl_name':'Kanwil DJP Riau','kpp':212,'kpp_name':'KPP Pratama Dumai'},
    {'kwl':40,'kwl_name':'Kanwil DJP Riau','kpp':213,'kpp_name':'KPP Pratama Rengat'},
    {'kwl':40,'kwl_name':'Kanwil DJP Riau','kpp':219,'kpp_name':'KPP Pratama Bengkalis'},
    {'kwl':40,'kwl_name':'Kanwil DJP Riau','kpp':221,'kpp_name':'KPP Pratama Bangkinang'},
    {'kwl':40,'kwl_name':'Kanwil DJP Riau','kpp':222,'kpp_name':'KPP Pratama Pangkalan Kerinci'},
    {'kwl':50,'kwl_name':'Kanwil DJP Sumatera Barat dan Jambi','kpp':202,'kpp_name':'KPP Pratama Bukittinggi'},
    {'kwl':50,'kwl_name':'Kanwil DJP Sumatera Barat dan Jambi','kpp':204,'kpp_name':'KPP Pratama Payakumbuh'},
    {'kwl':50,'kwl_name':'Kanwil DJP Sumatera Barat dan Jambi','kpp':203,'kpp_name':'KPP Pratama Solok'},
    {'kwl':50,'kwl_name':'Kanwil DJP Sumatera Barat dan Jambi','kpp':333,'kpp_name':'KPP Pratama Bangko'},
    {'kwl':50,'kwl_name':'Kanwil DJP Sumatera Barat dan Jambi','kpp':334,'kpp_name':'KPP Pratama Kuala Tungkal'},
    {'kwl':50,'kwl_name':'Kanwil DJP Sumatera Barat dan Jambi','kpp':332,'kpp_name':'KPP Pratama Muara Bungo'},
    {'kwl':50,'kwl_name':'Kanwil DJP Sumatera Barat dan Jambi','kpp':201,'kpp_name':'KPP Pratama Padang Satu'},
    {'kwl':50,'kwl_name':'Kanwil DJP Sumatera Barat dan Jambi','kpp':205,'kpp_name':'KPP Pratama Padang Dua'},
    {'kwl':50,'kwl_name':'Kanwil DJP Sumatera Barat dan Jambi','kpp':331,'kpp_name':'KPP Pratama Jambi Telanaipura'},
    {'kwl':50,'kwl_name':'Kanwil DJP Sumatera Barat dan Jambi','kpp':335,'kpp_name':'KPP Pratama Jambi Pelayangan'},
    {'kwl':60,'kwl_name':'Kanwil DJP Sumatera Selatan dan Kepulauan Bangka Belitung','kpp':308,'kpp_name':'KPP Madya Palembang'},
    {'kwl':60,'kwl_name':'Kanwil DJP Sumatera Selatan dan Kepulauan Bangka Belitung','kpp':302,'kpp_name':'KPP Pratama Baturaja'},
    {'kwl':60,'kwl_name':'Kanwil DJP Sumatera Selatan dan Kepulauan Bangka Belitung','kpp':309,'kpp_name':'KPP Pratama Lahat'},
    {'kwl':60,'kwl_name':'Kanwil DJP Sumatera Selatan dan Kepulauan Bangka Belitung','kpp':303,'kpp_name':'KPP Pratama Lubuk Linggau'},
    {'kwl':60,'kwl_name':'Kanwil DJP Sumatera Selatan dan Kepulauan Bangka Belitung','kpp':307,'kpp_name':'KPP Pratama Palembang Ilir Barat'},
    {'kwl':60,'kwl_name':'Kanwil DJP Sumatera Selatan dan Kepulauan Bangka Belitung','kpp':301,'kpp_name':'KPP Pratama Palembang Ilir Timur'},
    {'kwl':60,'kwl_name':'Kanwil DJP Sumatera Selatan dan Kepulauan Bangka Belitung','kpp':306,'kpp_name':'KPP Pratama Palembang Seberang Ulu'},
    {'kwl':60,'kwl_name':'Kanwil DJP Sumatera Selatan dan Kepulauan Bangka Belitung','kpp':312,'kpp_name':'KPP Pratama Kayu Agung'},
    {'kwl':60,'kwl_name':'Kanwil DJP Sumatera Selatan dan Kepulauan Bangka Belitung','kpp':313,'kpp_name':'KPP Pratama Prabumulih'},
    {'kwl':60,'kwl_name':'Kanwil DJP Sumatera Selatan dan Kepulauan Bangka Belitung','kpp':314,'kpp_name':'KPP Pratama Sekayu'},
    {'kwl':60,'kwl_name':'Kanwil DJP Sumatera Selatan dan Kepulauan Bangka Belitung','kpp':305,'kpp_name':'KPP Pratama Tanjung Pandan'},
    {'kwl':60,'kwl_name':'Kanwil DJP Sumatera Selatan dan Kepulauan Bangka Belitung','kpp':304,'kpp_name':'KPP Pratama Pangkal Pinang'},
    {'kwl':60,'kwl_name':'Kanwil DJP Sumatera Selatan dan Kepulauan Bangka Belitung','kpp':315,'kpp_name':'KPP Pratama Bangka'},
    {'kwl':70,'kwl_name':'Kanwil DJP Bengkulu dan Lampung','kpp':322,'kpp_name':'KPP Pratama Bandar Lampung Satu'},
    {'kwl':70,'kwl_name':'Kanwil DJP Bengkulu dan Lampung','kpp':323,'kpp_name':'KPP Pratama Bandar Lampung Dua'},
    {'kwl':70,'kwl_name':'Kanwil DJP Bengkulu dan Lampung','kpp':324,'kpp_name':'KPP Madya Bandar Lampung'},
    {'kwl':70,'kwl_name':'Kanwil DJP Bengkulu dan Lampung','kpp':325,'kpp_name':'KPP Pratama Natar'},
    {'kwl':70,'kwl_name':'Kanwil DJP Bengkulu dan Lampung','kpp':321,'kpp_name':'KPP Pratama Metro'},
    {'kwl':70,'kwl_name':'Kanwil DJP Bengkulu dan Lampung','kpp':326,'kpp_name':'KPP Pratama Kotabumi'},
    {'kwl':70,'kwl_name':'Kanwil DJP Bengkulu dan Lampung','kpp':311,'kpp_name':'KPP Pratama Bengkulu Dua'},
    {'kwl':70,'kwl_name':'Kanwil DJP Bengkulu dan Lampung','kpp':327,'kpp_name':'KPP Pratama Curup'},
    {'kwl':70,'kwl_name':'Kanwil DJP Bengkulu dan Lampung','kpp':328,'kpp_name':'KPP Pratama Bengkulu Satu'},
    {'kwl':80,'kwl_name':'Kanwil DJP Jakarta Pusat','kpp':73,'kpp_name':'KPP Madya Jakarta Pusat'},
    {'kwl':80,'kwl_name':'Kanwil DJP Jakarta Pusat','kpp':25,'kpp_name':'KPP Pratama Jakarta Gambir Satu'},
    {'kwl':80,'kwl_name':'Kanwil DJP Jakarta Pusat','kpp':28,'kpp_name':'KPP Pratama Jakarta Gambir Dua'},
    {'kwl':80,'kwl_name':'Kanwil DJP Jakarta Pusat','kpp':29,'kpp_name':'KPP Pratama Jakarta Gambir Tiga'},
    {'kwl':80,'kwl_name':'Kanwil DJP Jakarta Pusat','kpp':74,'kpp_name':'KPP Madya Dua Jakarta Pusat'},
    {'kwl':80,'kwl_name':'Kanwil DJP Jakarta Pusat','kpp':26,'kpp_name':'KPP Pratama Jakarta Sawah Besar Satu'},
    {'kwl':80,'kwl_name':'Kanwil DJP Jakarta Pusat','kpp':75,'kpp_name':'KPP Pratama Jakarta Sawah Besar Dua'},
    {'kwl':80,'kwl_name':'Kanwil DJP Jakarta Pusat','kpp':27,'kpp_name':'KPP Pratama Jakarta Kemayoran'},
    {'kwl':80,'kwl_name':'Kanwil DJP Jakarta Pusat','kpp':24,'kpp_name':'KPP Pratama Jakarta Cempaka Putih'},
    {'kwl':80,'kwl_name':'Kanwil DJP Jakarta Pusat','kpp':21,'kpp_name':'KPP Pratama Jakarta Menteng Satu'},
    {'kwl':80,'kwl_name':'Kanwil DJP Jakarta Pusat','kpp':71,'kpp_name':'KPP Pratama Jakarta Menteng Dua'},
    {'kwl':80,'kwl_name':'Kanwil DJP Jakarta Pusat','kpp':76,'kpp_name':'KPP Pratama Jakarta Menteng Tiga'},
    {'kwl':80,'kwl_name':'Kanwil DJP Jakarta Pusat','kpp':23,'kpp_name':'KPP Pratama Jakarta Senen'},
    {'kwl':80,'kwl_name':'Kanwil DJP Jakarta Pusat','kpp':22,'kpp_name':'KPP Pratama Jakarta Tanah Abang Satu'},
    {'kwl':80,'kwl_name':'Kanwil DJP Jakarta Pusat','kpp':72,'kpp_name':'KPP Pratama Jakarta Tanah Abang Dua'},
    {'kwl':80,'kwl_name':'Kanwil DJP Jakarta Pusat','kpp':77,'kpp_name':'KPP Pratama Jakarta Tanah Abang Tiga'},
    {'kwl':90,'kwl_name':'Kanwil DJP Jakarta Barat','kpp':38,'kpp_name':'KPP Madya Jakarta Barat'},
    {'kwl':90,'kwl_name':'Kanwil DJP Jakarta Barat','kpp':31,'kpp_name':'KPP Pratama Jakarta Palmerah'},
    {'kwl':90,'kwl_name':'Kanwil DJP Jakarta Barat','kpp':36,'kpp_name':'KPP Pratama Jakarta Grogol Petamburan'},
    {'kwl':90,'kwl_name':'Kanwil DJP Jakarta Barat','kpp':32,'kpp_name':'KPP Pratama Jakarta Tamansari'},
    {'kwl':90,'kwl_name':'Kanwil DJP Jakarta Barat','kpp':37,'kpp_name':'KPP Madya Dua Jakarta Barat'},
    {'kwl':90,'kwl_name':'Kanwil DJP Jakarta Barat','kpp':33,'kpp_name':'KPP Pratama Jakarta Tambora'},
    {'kwl':90,'kwl_name':'Kanwil DJP Jakarta Barat','kpp':34,'kpp_name':'KPP Pratama Jakarta Cengkareng'},
    {'kwl':90,'kwl_name':'Kanwil DJP Jakarta Barat','kpp':85,'kpp_name':'KPP Pratama Jakarta Kalideres'},
    {'kwl':90,'kwl_name':'Kanwil DJP Jakarta Barat','kpp':35,'kpp_name':'KPP Pratama Jakarta Kebon Jeruk Satu'},
    {'kwl':90,'kwl_name':'Kanwil DJP Jakarta Barat','kpp':39,'kpp_name':'KPP Pratama Jakarta Kebon Jeruk Dua'},
    {'kwl':90,'kwl_name':'Kanwil DJP Jakarta Barat','kpp':86,'kpp_name':'KPP Pratama Jakarta Kembangan'},
    {'kwl':100,'kwl_name':'Kanwil DJP Jakarta Selatan I','kpp':62,'kpp_name':'KPP Madya Jakarta Selatan I'},
    {'kwl':100,'kwl_name':'Kanwil DJP Jakarta Selatan I','kpp':11,'kpp_name':'KPP Pratama Jakarta Setiabudi Satu'},
    {'kwl':100,'kwl_name':'Kanwil DJP Jakarta Selatan I','kpp':18,'kpp_name':'KPP Pratama Jakarta Setiabudi Dua'},
    {'kwl':100,'kwl_name':'Kanwil DJP Jakarta Selatan I','kpp':15,'kpp_name':'KPP Pratama Jakarta Tebet'},
    {'kwl':100,'kwl_name':'Kanwil DJP Jakarta Selatan I','kpp':14,'kpp_name':'KPP Pratama Jakarta Mampang Prapatan'},
    {'kwl':100,'kwl_name':'Kanwil DJP Jakarta Selatan I','kpp':61,'kpp_name':'KPP Pratama Jakarta Pancoran'},
    {'kwl':100,'kwl_name':'Kanwil DJP Jakarta Selatan I','kpp':63,'kpp_name':'KPP Pratama Jakarta Setiabudi Tiga'},
    {'kwl':100,'kwl_name':'Kanwil DJP Jakarta Selatan I','kpp':67,'kpp_name':'KPP Madya Dua Jakarta Selatan I'},
    {'kwl':110,'kwl_name':'Kanwil DJP Jakarta Timur','kpp':7,'kpp_name':'KPP Madya Jakarta Timur'},
    {'kwl':110,'kwl_name':'Kanwil DJP Jakarta Timur','kpp':1,'kpp_name':'KPP Pratama Jakarta Matraman'},
    {'kwl':110,'kwl_name':'Kanwil DJP Jakarta Timur','kpp':2,'kpp_name':'KPP Pratama Jakarta Jatinegara'},
    {'kwl':110,'kwl_name':'Kanwil DJP Jakarta Timur','kpp':3,'kpp_name':'KPP Pratama Jakarta Pulogadung'},
    {'kwl':110,'kwl_name':'Kanwil DJP Jakarta Timur','kpp':4,'kpp_name':'KPP Pratama Jakarta Cakung'},
    {'kwl':110,'kwl_name':'Kanwil DJP Jakarta Timur','kpp':6,'kpp_name':'KPP Madya Dua Jakarta Timur'},
    {'kwl':110,'kwl_name':'Kanwil DJP Jakarta Timur','kpp':5,'kpp_name':'KPP Pratama Jakarta Kramat Jati'},
    {'kwl':110,'kwl_name':'Kanwil DJP Jakarta Timur','kpp':8,'kpp_name':'KPP Pratama Jakarta Duren Sawit'},
    {'kwl':110,'kwl_name':'Kanwil DJP Jakarta Timur','kpp':9,'kpp_name':'KPP Pratama Jakarta Pasar Rebo'},
    {'kwl':120,'kwl_name':'Kanwil DJP Jakarta Utara','kpp':46,'kpp_name':'KPP Madya Jakarta Utara'},
    {'kwl':120,'kwl_name':'Kanwil DJP Jakarta Utara','kpp':41,'kpp_name':'KPP Pratama Jakarta Penjaringan'},
    {'kwl':120,'kwl_name':'Kanwil DJP Jakarta Utara','kpp':44,'kpp_name':'KPP Pratama Jakarta Pademangan'},
    {'kwl':120,'kwl_name':'Kanwil DJP Jakarta Utara','kpp':42,'kpp_name':'KPP Pratama Jakarta Tanjung Priok'},
    {'kwl':120,'kwl_name':'Kanwil DJP Jakarta Utara','kpp':45,'kpp_name':'KPP Pratama Jakarta Koja'},
    {'kwl':120,'kwl_name':'Kanwil DJP Jakarta Utara','kpp':43,'kpp_name':'KPP Pratama Jakarta Kelapa Gading'},
    {'kwl':120,'kwl_name':'Kanwil DJP Jakarta Utara','kpp':48,'kpp_name':'KPP Madya Dua Jakarta Utara'},
    {'kwl':120,'kwl_name':'Kanwil DJP Jakarta Utara','kpp':47,'kpp_name':'KPP Pratama Jakarta Pluit'},
    {'kwl':130,'kwl_name':'Kanwil DJP Jakarta Khusus','kpp':52,'kpp_name':'KPP Penanaman Modal Asing Satu'},
    {'kwl':130,'kwl_name':'Kanwil DJP Jakarta Khusus','kpp':55,'kpp_name':'KPP Penanaman Modal Asing Dua'},
    {'kwl':130,'kwl_name':'Kanwil DJP Jakarta Khusus','kpp':56,'kpp_name':'KPP Penanaman Modal Asing Tiga'},
    {'kwl':130,'kwl_name':'Kanwil DJP Jakarta Khusus','kpp':57,'kpp_name':'KPP Penanaman Modal Asing Empat'},
    {'kwl':130,'kwl_name':'Kanwil DJP Jakarta Khusus','kpp':58,'kpp_name':'KPP Penanaman Modal Asing Lima'},
    {'kwl':130,'kwl_name':'Kanwil DJP Jakarta Khusus','kpp':59,'kpp_name':'KPP Penanaman Modal Asing Enam'},
    {'kwl':130,'kwl_name':'Kanwil DJP Jakarta Khusus','kpp':54,'kpp_name':'KPP Perusahaan Masuk Bursa'},
    {'kwl':130,'kwl_name':'Kanwil DJP Jakarta Khusus','kpp':53,'kpp_name':'KPP Badan dan Orang Asing'},
    {'kwl':130,'kwl_name':'Kanwil DJP Jakarta Khusus','kpp':81,'kpp_name':'KPP Minyak dan Gas Bumi'},
    {'kwl':140,'kwl_name':'Kanwil DJP Banten','kpp':415,'kpp_name':'KPP Madya Tangerang'},
    {'kwl':140,'kwl_name':'Kanwil DJP Banten','kpp':417,'kpp_name':'KPP Pratama Cilegon'},
    {'kwl':140,'kwl_name':'Kanwil DJP Banten','kpp':419,'kpp_name':'KPP Pratama Pandeglang'},
    {'kwl':140,'kwl_name':'Kanwil DJP Banten','kpp':411,'kpp_name':'KPP Pratama Serpong'},
    {'kwl':140,'kwl_name':'Kanwil DJP Banten','kpp':418,'kpp_name':'KPP Pratama Kosambi'},
    {'kwl':140,'kwl_name':'Kanwil DJP Banten','kpp':451,'kpp_name':'KPP Pratama Tigaraksa'},
    {'kwl':140,'kwl_name':'Kanwil DJP Banten','kpp':416,'kpp_name':'KPP Pratama Tangerang Timur'},
    {'kwl':140,'kwl_name':'Kanwil DJP Banten','kpp':402,'kpp_name':'KPP Pratama Tangerang Barat'},
    {'kwl':140,'kwl_name':'Kanwil DJP Banten','kpp':453,'kpp_name':'KPP Pratama Pondok Aren'},
    {'kwl':140,'kwl_name':'Kanwil DJP Banten','kpp':452,'kpp_name':'KPP Madya Dua Tangerang'},
    {'kwl':140,'kwl_name':'Kanwil DJP Banten','kpp':401,'kpp_name':'KPP Pratama Serang Barat'},
    {'kwl':140,'kwl_name':'Kanwil DJP Banten','kpp':454,'kpp_name':'KPP Pratama Serang Timur'},
    {'kwl':150,'kwl_name':'Kanwil DJP Jawa Barat I','kpp':441,'kpp_name':'KPP Madya Bandung'},
    {'kwl':150,'kwl_name':'Kanwil DJP Jawa Barat I','kpp':428,'kpp_name':'KPP Pratama Bandung Bojonagara'},
    {'kwl':150,'kwl_name':'Kanwil DJP Jawa Barat I','kpp':423,'kpp_name':'KPP Pratama Bandung Cibeunying'},
    {'kwl':150,'kwl_name':'Kanwil DJP Jawa Barat I','kpp':429,'kpp_name':'KPP Pratama Bandung Cicadas'},
    {'kwl':150,'kwl_name':'Kanwil DJP Jawa Barat I','kpp':424,'kpp_name':'KPP Madya Dua Bandung'},
    {'kwl':150,'kwl_name':'Kanwil DJP Jawa Barat I','kpp':422,'kpp_name':'KPP Pratama Bandung Tegallega'},
    {'kwl':150,'kwl_name':'Kanwil DJP Jawa Barat I','kpp':442,'kpp_name':'KPP Pratama Ciamis'},
    {'kwl':150,'kwl_name':'Kanwil DJP Jawa Barat I','kpp':406,'kpp_name':'KPP Pratama Cianjur'},
    {'kwl':150,'kwl_name':'Kanwil DJP Jawa Barat I','kpp':421,'kpp_name':'KPP Pratama Cimahi'},
    {'kwl':150,'kwl_name':'Kanwil DJP Jawa Barat I','kpp':409,'kpp_name':'KPP Pratama Purwakarta'},
    {'kwl':150,'kwl_name':'Kanwil DJP Jawa Barat I','kpp':445,'kpp_name':'KPP Pratama Soreang'},
    {'kwl':150,'kwl_name':'Kanwil DJP Jawa Barat I','kpp':444,'kpp_name':'KPP Pratama Majalaya'},
    {'kwl':150,'kwl_name':'Kanwil DJP Jawa Barat I','kpp':405,'kpp_name':'KPP Pratama Sukabumi'},
    {'kwl':150,'kwl_name':'Kanwil DJP Jawa Barat I','kpp':446,'kpp_name':'KPP Pratama Sumedang'},
    {'kwl':150,'kwl_name':'Kanwil DJP Jawa Barat I','kpp':425,'kpp_name':'KPP Pratama Tasikmalaya'},
    {'kwl':150,'kwl_name':'Kanwil DJP Jawa Barat I','kpp':443,'kpp_name':'KPP Pratama Garut'},
    {'kwl':160,'kwl_name':'Kanwil DJP Jawa Barat II','kpp':431,'kpp_name':'KPP Madya Bekasi'},
    {'kwl':160,'kwl_name':'Kanwil DJP Jawa Barat II','kpp':435,'kpp_name':'KPP Pratama Cibitung'},
    {'kwl':160,'kwl_name':'Kanwil DJP Jawa Barat II','kpp':413,'kpp_name':'KPP Pratama Cikarang Selatan'},
    {'kwl':160,'kwl_name':'Kanwil DJP Jawa Barat II','kpp':414,'kpp_name':'KPP Pratama Cikarang Utara'},
    {'kwl':160,'kwl_name':'Kanwil DJP Jawa Barat II','kpp':438,'kpp_name':'KPP Pratama Kuningan'},
    {'kwl':160,'kwl_name':'Kanwil DJP Jawa Barat II','kpp':437,'kpp_name':'KPP Pratama Indramayu'},
    {'kwl':160,'kwl_name':'Kanwil DJP Jawa Barat II','kpp':408,'kpp_name':'KPP Pratama Karawang'},
    {'kwl':160,'kwl_name':'Kanwil DJP Jawa Barat II','kpp':433,'kpp_name':'KPP Madya Karawang'},
    {'kwl':160,'kwl_name':'Kanwil DJP Jawa Barat II','kpp':439,'kpp_name':'KPP Pratama Subang'},
    {'kwl':160,'kwl_name':'Kanwil DJP Jawa Barat II','kpp':426,'kpp_name':'KPP Pratama Cirebon Satu'},
    {'kwl':160,'kwl_name':'Kanwil DJP Jawa Barat II','kpp':455,'kpp_name':'KPP Pratama Cirebon Dua'},
    {'kwl':170,'kwl_name':'Kanwil DJP Jawa Tengah I','kpp':511,'kpp_name':'KPP Madya Semarang'},
    {'kwl':170,'kwl_name':'Kanwil DJP Jawa Tengah I','kpp':514,'kpp_name':'KPP Pratama Blora'},
    {'kwl':170,'kwl_name':'Kanwil DJP Jawa Tengah I','kpp':515,'kpp_name':'KPP Pratama Demak'},
    {'kwl':170,'kwl_name':'Kanwil DJP Jawa Tengah I','kpp':506,'kpp_name':'KPP Pratama Kudus'},
    {'kwl':170,'kwl_name':'Kanwil DJP Jawa Tengah I','kpp':516,'kpp_name':'KPP Pratama Jepara'},
    {'kwl':170,'kwl_name':'Kanwil DJP Jawa Tengah I','kpp':507,'kpp_name':'KPP Pratama Pati'},
    {'kwl':170,'kwl_name':'Kanwil DJP Jawa Tengah I','kpp':502,'kpp_name':'KPP Pratama Pekalongan'},
    {'kwl':170,'kwl_name':'Kanwil DJP Jawa Tengah I','kpp':513,'kpp_name':'KPP Pratama Batang'},
    {'kwl':170,'kwl_name':'Kanwil DJP Jawa Tengah I','kpp':505,'kpp_name':'KPP Pratama Salatiga'},
    {'kwl':170,'kwl_name':'Kanwil DJP Jawa Tengah I','kpp':503,'kpp_name':'KPP Pratama Semarang Barat'},
    {'kwl':170,'kwl_name':'Kanwil DJP Jawa Tengah I','kpp':517,'kpp_name':'KPP Pratama Semarang Candisari'},
    {'kwl':170,'kwl_name':'Kanwil DJP Jawa Tengah I','kpp':508,'kpp_name':'KPP Pratama Semarang Selatan'},
    {'kwl':170,'kwl_name':'Kanwil DJP Jawa Tengah I','kpp':512,'kpp_name':'KPP Madya Dua Semarang'},
    {'kwl':170,'kwl_name':'Kanwil DJP Jawa Tengah I','kpp':509,'kpp_name':'KPP Pratama Semarang Tengah'},
    {'kwl':170,'kwl_name':'Kanwil DJP Jawa Tengah I','kpp':504,'kpp_name':'KPP Pratama Semarang Timur'},
    {'kwl':170,'kwl_name':'Kanwil DJP Jawa Tengah I','kpp':518,'kpp_name':'KPP Pratama Semarang Gayamsari'},
    {'kwl':170,'kwl_name':'Kanwil DJP Jawa Tengah I','kpp':501,'kpp_name':'KPP Pratama Tegal'},
    {'kwl':180,'kwl_name':'Kanwil DJP Jawa Tengah II','kpp':521,'kpp_name':'KPP Pratama Purwokerto'},
    {'kwl':180,'kwl_name':'Kanwil DJP Jawa Tengah II','kpp':526,'kpp_name':'KPP Pratama Surakarta'},
    {'kwl':180,'kwl_name':'Kanwil DJP Jawa Tengah II','kpp':529,'kpp_name':'KPP Pratama Purbalingga'},
    {'kwl':180,'kwl_name':'Kanwil DJP Jawa Tengah II','kpp':527,'kpp_name':'KPP Pratama Boyolali'},
    {'kwl':180,'kwl_name':'Kanwil DJP Jawa Tengah II','kpp':533,'kpp_name':'KPP Pratama Temanggung'},
    {'kwl':180,'kwl_name':'Kanwil DJP Jawa Tengah II','kpp':531,'kpp_name':'KPP Pratama Purworejo'},
    {'kwl':180,'kwl_name':'Kanwil DJP Jawa Tengah II','kpp':525,'kpp_name':'KPP Pratama Klaten'},
    {'kwl':180,'kwl_name':'Kanwil DJP Jawa Tengah II','kpp':522,'kpp_name':'KPP Pratama Cilacap'},
    {'kwl':180,'kwl_name':'Kanwil DJP Jawa Tengah II','kpp':528,'kpp_name':'KPP Pratama Karanganyar'},
    {'kwl':180,'kwl_name':'Kanwil DJP Jawa Tengah II','kpp':523,'kpp_name':'KPP Pratama Kebumen'},
    {'kwl':180,'kwl_name':'Kanwil DJP Jawa Tengah II','kpp':532,'kpp_name':'KPP Pratama Sukoharjo'},
    {'kwl':180,'kwl_name':'Kanwil DJP Jawa Tengah II','kpp':524,'kpp_name':'KPP Pratama Magelang'},
    {'kwl':190,'kwl_name':'Kanwil DJP Daerah Istimewa Yogyakarta','kpp':542,'kpp_name':'KPP Pratama Sleman'},
    {'kwl':190,'kwl_name':'Kanwil DJP Daerah Istimewa Yogyakarta','kpp':541,'kpp_name':'KPP Pratama Yogyakarta'},
    {'kwl':190,'kwl_name':'Kanwil DJP Daerah Istimewa Yogyakarta','kpp':545,'kpp_name':'KPP Pratama Wonosari'},
    {'kwl':190,'kwl_name':'Kanwil DJP Daerah Istimewa Yogyakarta','kpp':544,'kpp_name':'KPP Pratama Wates'},
    {'kwl':190,'kwl_name':'Kanwil DJP Daerah Istimewa Yogyakarta','kpp':543,'kpp_name':'KPP Pratama Bantul'},
    {'kwl':200,'kwl_name':'Kanwil DJP Jawa Timur I','kpp':631,'kpp_name':'KPP Madya Surabaya'},
    {'kwl':200,'kwl_name':'Kanwil DJP Jawa Timur I','kpp':611,'kpp_name':'KPP Pratama Surabaya Genteng'},
    {'kwl':200,'kwl_name':'Kanwil DJP Jawa Timur I','kpp':606,'kpp_name':'KPP Pratama Surabaya Gubeng'},
    {'kwl':200,'kwl_name':'Kanwil DJP Jawa Timur I','kpp':605,'kpp_name':'KPP Pratama Surabaya Krembangan'},
    {'kwl':200,'kwl_name':'Kanwil DJP Jawa Timur I','kpp':613,'kpp_name':'KPP Pratama Surabaya Pabean Cantikan'},
    {'kwl':200,'kwl_name':'Kanwil DJP Jawa Timur I','kpp':615,'kpp_name':'KPP Pratama Surabaya Rungkut'},
    {'kwl':200,'kwl_name':'Kanwil DJP Jawa Timur I','kpp':614,'kpp_name':'KPP Pratama Surabaya Sawahan'},
    {'kwl':200,'kwl_name':'Kanwil DJP Jawa Timur I','kpp':619,'kpp_name':'KPP Pratama Surabaya Mulyorejo'},
    {'kwl':200,'kwl_name':'Kanwil DJP Jawa Timur I','kpp':616,'kpp_name':'KPP Madya Dua Surabaya'},
    {'kwl':200,'kwl_name':'Kanwil DJP Jawa Timur I','kpp':604,'kpp_name':'KPP Pratama Surabaya Sukomanunggal'},
    {'kwl':200,'kwl_name':'Kanwil DJP Jawa Timur I','kpp':607,'kpp_name':'KPP Pratama Surabaya Tegalsari'},
    {'kwl':200,'kwl_name':'Kanwil DJP Jawa Timur I','kpp':609,'kpp_name':'KPP Pratama Surabaya Wonocolo'},
    {'kwl':200,'kwl_name':'Kanwil DJP Jawa Timur I','kpp':618,'kpp_name':'KPP Pratama Surabaya Karangpilang'},
    {'kwl':210,'kwl_name':'Kanwil DJP Jawa Timur II','kpp':641,'kpp_name':'KPP Madya Sidoarjo'},
    {'kwl':210,'kwl_name':'Kanwil DJP Jawa Timur II','kpp':601,'kpp_name':'KPP Pratama Bojonegoro'},
    {'kwl':210,'kwl_name':'Kanwil DJP Jawa Timur II','kpp':645,'kpp_name':'KPP Pratama Lamongan'},
    {'kwl':210,'kwl_name':'Kanwil DJP Jawa Timur II','kpp':612,'kpp_name':'KPP Madya Gresik'},
    {'kwl':210,'kwl_name':'Kanwil DJP Jawa Timur II','kpp':642,'kpp_name':'KPP Pratama Gresik'},
    {'kwl':210,'kwl_name':'Kanwil DJP Jawa Timur II','kpp':621,'kpp_name':'KPP Pratama Madiun'},
    {'kwl':210,'kwl_name':'Kanwil DJP Jawa Timur II','kpp':602,'kpp_name':'KPP Pratama Mojokerto'},
    {'kwl':210,'kwl_name':'Kanwil DJP Jawa Timur II','kpp':603,'kpp_name':'KPP Pratama Sidoarjo Barat'},
    {'kwl':210,'kwl_name':'Kanwil DJP Jawa Timur II','kpp':617,'kpp_name':'KPP Pratama Sidoarjo Selatan'},
    {'kwl':210,'kwl_name':'Kanwil DJP Jawa Timur II','kpp':643,'kpp_name':'KPP Pratama Sidoarjo Utara'},
    {'kwl':210,'kwl_name':'Kanwil DJP Jawa Timur II','kpp':648,'kpp_name':'KPP Pratama Tuban'},
    {'kwl':210,'kwl_name':'Kanwil DJP Jawa Timur II','kpp':646,'kpp_name':'KPP Pratama Ngawi'},
    {'kwl':210,'kwl_name':'Kanwil DJP Jawa Timur II','kpp':647,'kpp_name':'KPP Pratama Ponorogo'},
    {'kwl':210,'kwl_name':'Kanwil DJP Jawa Timur II','kpp':608,'kpp_name':'KPP Pratama Pamekasan'},
    {'kwl':210,'kwl_name':'Kanwil DJP Jawa Timur II','kpp':644,'kpp_name':'KPP Pratama Bangkalan'},
    {'kwl':210,'kwl_name':'Kanwil DJP Jawa Timur II','kpp':649,'kpp_name':'KPP Pratama Jombang'},
    {'kwl':220,'kwl_name':'Kanwil DJP Jawa Timur III','kpp':651,'kpp_name':'KPP Madya Malang'},
    {'kwl':220,'kwl_name':'Kanwil DJP Jawa Timur III','kpp':627,'kpp_name':'KPP Pratama Banyuwangi'},
    {'kwl':220,'kwl_name':'Kanwil DJP Jawa Timur III','kpp':628,'kpp_name':'KPP Pratama Batu'},
    {'kwl':220,'kwl_name':'Kanwil DJP Jawa Timur III','kpp':657,'kpp_name':'KPP Pratama Singosari'},
    {'kwl':220,'kwl_name':'Kanwil DJP Jawa Timur III','kpp':654,'kpp_name':'KPP Pratama Kepanjen'},
    {'kwl':220,'kwl_name':'Kanwil DJP Jawa Timur III','kpp':653,'kpp_name':'KPP Pratama Blitar'},
    {'kwl':220,'kwl_name':'Kanwil DJP Jawa Timur III','kpp':626,'kpp_name':'KPP Pratama Jember'},
    {'kwl':220,'kwl_name':'Kanwil DJP Jawa Timur III','kpp':622,'kpp_name':'KPP Pratama Kediri'},
    {'kwl':220,'kwl_name':'Kanwil DJP Jawa Timur III','kpp':655,'kpp_name':'KPP Pratama Pare'},
    {'kwl':220,'kwl_name':'Kanwil DJP Jawa Timur III','kpp':623,'kpp_name':'KPP Pratama Malang Selatan'},
    {'kwl':220,'kwl_name':'Kanwil DJP Jawa Timur III','kpp':652,'kpp_name':'KPP Pratama Malang Utara'},
    {'kwl':220,'kwl_name':'Kanwil DJP Jawa Timur III','kpp':624,'kpp_name':'KPP Pratama Pasuruan'},
    {'kwl':220,'kwl_name':'Kanwil DJP Jawa Timur III','kpp':625,'kpp_name':'KPP Pratama Probolinggo'},
    {'kwl':220,'kwl_name':'Kanwil DJP Jawa Timur III','kpp':656,'kpp_name':'KPP Pratama Situbondo'},
    {'kwl':220,'kwl_name':'Kanwil DJP Jawa Timur III','kpp':629,'kpp_name':'KPP Pratama Tulungagung'},
    {'kwl':230,'kwl_name':'Kanwil DJP Kalimantan Barat','kpp':703,'kpp_name':'KPP Pratama Ketapang'},
    {'kwl':230,'kwl_name':'Kanwil DJP Kalimantan Barat','kpp':704,'kpp_name':'KPP Pratama Kubu Raya'},
    {'kwl':230,'kwl_name':'Kanwil DJP Kalimantan Barat','kpp':705,'kpp_name':'KPP Pratama Sanggau'},
    {'kwl':230,'kwl_name':'Kanwil DJP Kalimantan Barat','kpp':702,'kpp_name':'KPP Pratama Singkawang'},
    {'kwl':230,'kwl_name':'Kanwil DJP Kalimantan Barat','kpp':706,'kpp_name':'KPP Pratama Sintang'},
    {'kwl':230,'kwl_name':'Kanwil DJP Kalimantan Barat','kpp':701,'kpp_name':'KPP Pratama Pontianak Barat'},
    {'kwl':230,'kwl_name':'Kanwil DJP Kalimantan Barat','kpp':707,'kpp_name':'KPP Pratama Pontianak Timur'},
    {'kwl':240,'kwl_name':'Kanwil DJP Kalimantan Selatan dan Tengah','kpp':732,'kpp_name':'KPP Pratama Banjarbaru'},
    {'kwl':240,'kwl_name':'Kanwil DJP Kalimantan Selatan dan Tengah','kpp':733,'kpp_name':'KPP Pratama Barabai'},
    {'kwl':240,'kwl_name':'Kanwil DJP Kalimantan Selatan dan Tengah','kpp':734,'kpp_name':'KPP Pratama Batulicin'},
    {'kwl':240,'kwl_name':'Kanwil DJP Kalimantan Selatan dan Tengah','kpp':711,'kpp_name':'KPP Pratama Palangkaraya'},
    {'kwl':240,'kwl_name':'Kanwil DJP Kalimantan Selatan dan Tengah','kpp':713,'kpp_name':'KPP Pratama Pangkalan Bun'},
    {'kwl':240,'kwl_name':'Kanwil DJP Kalimantan Selatan dan Tengah','kpp':735,'kpp_name':'KPP Pratama Tanjung'},
    {'kwl':240,'kwl_name':'Kanwil DJP Kalimantan Selatan dan Tengah','kpp':712,'kpp_name':'KPP Pratama Sampit'},
    {'kwl':240,'kwl_name':'Kanwil DJP Kalimantan Selatan dan Tengah','kpp':714,'kpp_name':'KPP Pratama Muara Teweh'},
    {'kwl':240,'kwl_name':'Kanwil DJP Kalimantan Selatan dan Tengah','kpp':736,'kpp_name':'KPP Madya Banjarmasin'},
    {'kwl':240,'kwl_name':'Kanwil DJP Kalimantan Selatan dan Tengah','kpp':731,'kpp_name':'KPP Pratama Banjarmasin'},
    {'kwl':250,'kwl_name':'Kanwil DJP Kalimantan Timur dan Utara','kpp':725,'kpp_name':'KPP Madya Balikpapan'},
    {'kwl':250,'kwl_name':'Kanwil DJP Kalimantan Timur dan Utara','kpp':726,'kpp_name':'KPP Pratama Penajam'},
    {'kwl':250,'kwl_name':'Kanwil DJP Kalimantan Timur dan Utara','kpp':724,'kpp_name':'KPP Pratama Bontang'},
    {'kwl':250,'kwl_name':'Kanwil DJP Kalimantan Timur dan Utara','kpp':723,'kpp_name':'KPP Pratama Tarakan'},
    {'kwl':250,'kwl_name':'Kanwil DJP Kalimantan Timur dan Utara','kpp':727,'kpp_name':'KPP Pratama Tanjung Redeb'},
    {'kwl':250,'kwl_name':'Kanwil DJP Kalimantan Timur dan Utara','kpp':728,'kpp_name':'KPP Pratama Tenggarong'},
    {'kwl':250,'kwl_name':'Kanwil DJP Kalimantan Timur dan Utara','kpp':729,'kpp_name':'KPP Pratama Balikpapan Barat'},
    {'kwl':250,'kwl_name':'Kanwil DJP Kalimantan Timur dan Utara','kpp':721,'kpp_name':'KPP Pratama Balikpapan Timur'},
    {'kwl':250,'kwl_name':'Kanwil DJP Kalimantan Timur dan Utara','kpp':722,'kpp_name':'KPP Pratama Samarinda Ilir'},
    {'kwl':250,'kwl_name':'Kanwil DJP Kalimantan Timur dan Utara','kpp':741,'kpp_name':'KPP Pratama Samarinda Ulu'},
    {'kwl':260,'kwl_name':'Kanwil DJP Sulawesi Selatan, Barat, dan Tenggara','kpp':812,'kpp_name':'KPP Madya Makassar'},
    {'kwl':260,'kwl_name':'Kanwil DJP Sulawesi Selatan, Barat, dan Tenggara','kpp':801,'kpp_name':'KPP Pratama Makassar Utara'},
    {'kwl':260,'kwl_name':'Kanwil DJP Sulawesi Selatan, Barat, dan Tenggara','kpp':804,'kpp_name':'KPP Pratama Makassar Barat'},
    {'kwl':260,'kwl_name':'Kanwil DJP Sulawesi Selatan, Barat, dan Tenggara','kpp':805,'kpp_name':'KPP Pratama Makassar Selatan'},
    {'kwl':260,'kwl_name':'Kanwil DJP Sulawesi Selatan, Barat, dan Tenggara','kpp':806,'kpp_name':'KPP Pratama Bulukumba'},
    {'kwl':260,'kwl_name':'Kanwil DJP Sulawesi Selatan, Barat, dan Tenggara','kpp':807,'kpp_name':'KPP Pratama Bantaeng'},
    {'kwl':260,'kwl_name':'Kanwil DJP Sulawesi Selatan, Barat, dan Tenggara','kpp':808,'kpp_name':'KPP Pratama Watampone'},
    {'kwl':260,'kwl_name':'Kanwil DJP Sulawesi Selatan, Barat, dan Tenggara','kpp':809,'kpp_name':'KPP Pratama Maros'},
    {'kwl':260,'kwl_name':'Kanwil DJP Sulawesi Selatan, Barat, dan Tenggara','kpp':803,'kpp_name':'KPP Pratama Palopo'},
    {'kwl':260,'kwl_name':'Kanwil DJP Sulawesi Selatan, Barat, dan Tenggara','kpp':802,'kpp_name':'KPP Pratama Parepare'},
    {'kwl':260,'kwl_name':'Kanwil DJP Sulawesi Selatan, Barat, dan Tenggara','kpp':813,'kpp_name':'KPP Pratama Majene'},
    {'kwl':260,'kwl_name':'Kanwil DJP Sulawesi Selatan, Barat, dan Tenggara','kpp':814,'kpp_name':'KPP Pratama Mamuju'},
    {'kwl':260,'kwl_name':'Kanwil DJP Sulawesi Selatan, Barat, dan Tenggara','kpp':811,'kpp_name':'KPP Pratama Kendari'},
    {'kwl':260,'kwl_name':'Kanwil DJP Sulawesi Selatan, Barat, dan Tenggara','kpp':815,'kpp_name':'KPP Pratama Kolaka'},
    {'kwl':260,'kwl_name':'Kanwil DJP Sulawesi Selatan, Barat, dan Tenggara','kpp':816,'kpp_name':'KPP Pratama Baubau'},
    {'kwl':270,'kwl_name':'Kanwil DJP Sulawesi Utara, Tengah, Gorontalo, dan Maluku Utara','kpp':823,'kpp_name':'KPP Pratama Bitung'},
    {'kwl':270,'kwl_name':'Kanwil DJP Sulawesi Utara, Tengah, Gorontalo, dan Maluku Utara','kpp':821,'kpp_name':'KPP Pratama Manado'},
    {'kwl':270,'kwl_name':'Kanwil DJP Sulawesi Utara, Tengah, Gorontalo, dan Maluku Utara','kpp':822,'kpp_name':'KPP Pratama Gorontalo'},
    {'kwl':270,'kwl_name':'Kanwil DJP Sulawesi Utara, Tengah, Gorontalo, dan Maluku Utara','kpp':824,'kpp_name':'KPP Pratama Kotamobagu'},
    {'kwl':270,'kwl_name':'Kanwil DJP Sulawesi Utara, Tengah, Gorontalo, dan Maluku Utara','kpp':825,'kpp_name':'KPP Pratama Tahuna'},
    {'kwl':270,'kwl_name':'Kanwil DJP Sulawesi Utara, Tengah, Gorontalo, dan Maluku Utara','kpp':832,'kpp_name':'KPP Pratama Luwuk'},
    {'kwl':270,'kwl_name':'Kanwil DJP Sulawesi Utara, Tengah, Gorontalo, dan Maluku Utara','kpp':831,'kpp_name':'KPP Pratama Palu'},
    {'kwl':270,'kwl_name':'Kanwil DJP Sulawesi Utara, Tengah, Gorontalo, dan Maluku Utara','kpp':833,'kpp_name':'KPP Pratama Poso'},
    {'kwl':270,'kwl_name':'Kanwil DJP Sulawesi Utara, Tengah, Gorontalo, dan Maluku Utara','kpp':834,'kpp_name':'KPP Pratama Toli Toli'},
    {'kwl':270,'kwl_name':'Kanwil DJP Sulawesi Utara, Tengah, Gorontalo, dan Maluku Utara','kpp':943,'kpp_name':'KPP Pratama Tobelo'},
    {'kwl':270,'kwl_name':'Kanwil DJP Sulawesi Utara, Tengah, Gorontalo, dan Maluku Utara','kpp':942,'kpp_name':'KPP Pratama Ternate'},
    {'kwl':280,'kwl_name':'Kanwil DJP Bali','kpp':904,'kpp_name':'KPP Madya Denpasar'},
    {'kwl':280,'kwl_name':'Kanwil DJP Bali','kpp':905,'kpp_name':'KPP Pratama Badung Selatan'},
    {'kwl':280,'kwl_name':'Kanwil DJP Bali','kpp':906,'kpp_name':'KPP Pratama Badung Utara'},
    {'kwl':280,'kwl_name':'Kanwil DJP Bali','kpp':907,'kpp_name':'KPP Pratama Gianyar'},
    {'kwl':280,'kwl_name':'Kanwil DJP Bali','kpp':901,'kpp_name':'KPP Pratama Denpasar Barat'},
    {'kwl':280,'kwl_name':'Kanwil DJP Bali','kpp':903,'kpp_name':'KPP Pratama Denpasar Timur'},
    {'kwl':280,'kwl_name':'Kanwil DJP Bali','kpp':902,'kpp_name':'KPP Pratama Singaraja'},
    {'kwl':280,'kwl_name':'Kanwil DJP Bali','kpp':908,'kpp_name':'KPP Pratama Tabanan'},
    {'kwl':290,'kwl_name':'Kanwil DJP Nusa Tenggara','kpp':914,'kpp_name':'KPP Pratama Mataram Timur'},
    {'kwl':290,'kwl_name':'Kanwil DJP Nusa Tenggara','kpp':911,'kpp_name':'KPP Pratama Mataram Barat'},
    {'kwl':290,'kwl_name':'Kanwil DJP Nusa Tenggara','kpp':915,'kpp_name':'KPP Pratama Praya'},
    {'kwl':290,'kwl_name':'Kanwil DJP Nusa Tenggara','kpp':913,'kpp_name':'KPP Pratama Sumbawa Besar'},
    {'kwl':290,'kwl_name':'Kanwil DJP Nusa Tenggara','kpp':912,'kpp_name':'KPP Pratama Raba Bima'},
    {'kwl':290,'kwl_name':'Kanwil DJP Nusa Tenggara','kpp':923,'kpp_name':'KPP Pratama Ende'},
    {'kwl':290,'kwl_name':'Kanwil DJP Nusa Tenggara','kpp':924,'kpp_name':'KPP Pratama Ruteng'},
    {'kwl':290,'kwl_name':'Kanwil DJP Nusa Tenggara','kpp':921,'kpp_name':'KPP Pratama Maumere'},
    {'kwl':290,'kwl_name':'Kanwil DJP Nusa Tenggara','kpp':925,'kpp_name':'KPP Pratama Atambua'},
    {'kwl':290,'kwl_name':'Kanwil DJP Nusa Tenggara','kpp':922,'kpp_name':'KPP Pratama Kupang'},
    {'kwl':290,'kwl_name':'Kanwil DJP Nusa Tenggara','kpp':926,'kpp_name':'KPP Pratama Waingapu'},
    {'kwl':300,'kwl_name':'Kanwil DJP Papua dan Maluku','kpp':941,'kpp_name':'KPP Pratama Ambon'},
    {'kwl':300,'kwl_name':'Kanwil DJP Papua dan Maluku','kpp':952,'kpp_name':'KPP Pratama Jayapura'},
    {'kwl':300,'kwl_name':'Kanwil DJP Papua dan Maluku','kpp':956,'kpp_name':'KPP Pratama Merauke'},
    {'kwl':300,'kwl_name':'Kanwil DJP Papua dan Maluku','kpp':955,'kpp_name':'KPP Pratama Manokwari'},
    {'kwl':300,'kwl_name':'Kanwil DJP Papua dan Maluku','kpp':951,'kpp_name':'KPP Pratama Sorong'},
    {'kwl':300,'kwl_name':'Kanwil DJP Papua dan Maluku','kpp':953,'kpp_name':'KPP Pratama Timika'},
    {'kwl':300,'kwl_name':'Kanwil DJP Papua dan Maluku','kpp':954,'kpp_name':'KPP Pratama Biak'},
    {'kwl':310,'kwl_name':'Kanwil DJP Wajib Pajak Besar','kpp':91,'kpp_name':'KPP Wajib Pajak Besar Satu'},
    {'kwl':310,'kwl_name':'Kanwil DJP Wajib Pajak Besar','kpp':92,'kpp_name':'KPP Wajib Pajak Besar Dua'},
    {'kwl':310,'kwl_name':'Kanwil DJP Wajib Pajak Besar','kpp':51,'kpp_name':'KPP Wajib Pajak Besar Tiga'},
    {'kwl':310,'kwl_name':'Kanwil DJP Wajib Pajak Besar','kpp':93,'kpp_name':'KPP Wajib Pajak Besar Empat'},
    {'kwl':320,'kwl_name':'Kanwil DJP Jakarta Selatan II','kpp':12,'kpp_name':'KPP Pratama Jakarta Kebayoran Baru Satu'},
    {'kwl':320,'kwl_name':'Kanwil DJP Jakarta Selatan II','kpp':19,'kpp_name':'KPP Pratama Jakarta Kebayoran Baru Dua'},
    {'kwl':320,'kwl_name':'Kanwil DJP Jakarta Selatan II','kpp':64,'kpp_name':'KPP Madya Dua Jakarta Selatan II'},
    {'kwl':320,'kwl_name':'Kanwil DJP Jakarta Selatan II','kpp':65,'kpp_name':'KPP Madya Jakarta Selatan II'},
    {'kwl':320,'kwl_name':'Kanwil DJP Jakarta Selatan II','kpp':13,'kpp_name':'KPP Pratama Jakarta Kebayoran Lama'},
    {'kwl':320,'kwl_name':'Kanwil DJP Jakarta Selatan II','kpp':66,'kpp_name':'KPP Pratama Jakarta Pesanggrahan'},
    {'kwl':320,'kwl_name':'Kanwil DJP Jakarta Selatan II','kpp':17,'kpp_name':'KPP Pratama Jakarta Pasar Minggu'},
    {'kwl':320,'kwl_name':'Kanwil DJP Jakarta Selatan II','kpp':16,'kpp_name':'KPP Pratama Jakarta Cilandak'},
    {'kwl':320,'kwl_name':'Kanwil DJP Jakarta Selatan II','kpp':68,'kpp_name':'KPP Pratama Jakarta Jagakarsa'},
    {'kwl':330,'kwl_name':'Kanwil DJP Jawa Barat III','kpp':447,'kpp_name':'KPP Pratama Pondok Gede'},
    {'kwl':330,'kwl_name':'Kanwil DJP Jawa Barat III','kpp':427,'kpp_name':'KPP Pratama Bekasi Barat'},
    {'kwl':330,'kwl_name':'Kanwil DJP Jawa Barat III','kpp':432,'kpp_name':'KPP Madya Kota Bekasi'},
    {'kwl':330,'kwl_name':'Kanwil DJP Jawa Barat III','kpp':407,'kpp_name':'KPP Pratama Bekasi Utara'},
    {'kwl':330,'kwl_name':'Kanwil DJP Jawa Barat III','kpp':448,'kpp_name':'KPP Pratama Depok Sawangan'},
    {'kwl':330,'kwl_name':'Kanwil DJP Jawa Barat III','kpp':412,'kpp_name':'KPP Pratama Depok Cimanggis'},
    {'kwl':330,'kwl_name':'Kanwil DJP Jawa Barat III','kpp':403,'kpp_name':'KPP Pratama Cibinong'},
    {'kwl':330,'kwl_name':'Kanwil DJP Jawa Barat III','kpp':434,'kpp_name':'KPP Pratama Ciawi'},
    {'kwl':330,'kwl_name':'Kanwil DJP Jawa Barat III','kpp':436,'kpp_name':'KPP Pratama Cileungsi'},
    {'kwl':330,'kwl_name':'Kanwil DJP Jawa Barat III','kpp':404,'kpp_name':'KPP Pratama Bogor'},
    {'kwl':330,'kwl_name':'Kanwil DJP Jawa Barat III','kpp':449,'kpp_name':'KPP Madya Bogor'},
    {'kwl':340,'kwl_name':'Kanwil DJP Kepulauan Riau','kpp':217,'kpp_name':'KPP Madya Batam'},
    {'kwl':340,'kwl_name':'Kanwil DJP Kepulauan Riau','kpp':214,'kpp_name':'KPP Pratama Tanjung Pinang'},
    {'kwl':340,'kwl_name':'Kanwil DJP Kepulauan Riau','kpp':215,'kpp_name':'KPP Pratama Batam Utara'},
    {'kwl':340,'kwl_name':'Kanwil DJP Kepulauan Riau','kpp':225,'kpp_name':'KPP Pratama Batam Selatan'},
    {'kwl':340,'kwl_name':'Kanwil DJP Kepulauan Riau','kpp':223,'kpp_name':'KPP Pratama Tanjung Balai Karimun'},
    {'kwl':340,'kwl_name':'Kanwil DJP Kepulauan Riau','kpp':224,'kpp_name':'KPP Pratama Bintan'},
]

# =============================================================
# PETA CONSTANTS
# =============================================================
# Kolom wajib — dicari berdasarkan NAMA (bukan posisi)
# Penambahan kolom baru di Excel tidak akan mengganggu pembacaan
PETA_REQUIRED_COLS = [
    'NAMA_KPP', 'KANWIL', 'PENERBIT', 'JENIS_LHA', 'NM_KATEGORI',
    'POTENSI_LHA', 'POTENSI_AWAL', 'POTENSI_AKHIR',
    'NILAI_PEMBAYARAN', 'NO_SP2DK', 'NO_LHP2DK',
]

# Koreksi nama KPP — perbedaan antara kolom NAMA_KPP di LHA vs data koordinat
KPP_NAME_FIX = {
    'KPP PRATAMA MATARAM BARAT': 'KPP PRATAMA MATARAM TIMUR',
    'KPP PRATAMA PANGKALANBUN':  'KPP PRATAMA PANGKALAN BUN',
    'KPP PRATAMA TAPAK TUAN':    'KPP PRATAMA TAPAKTUAN',
}

# Urutan KLU chip di sidebar (berdasarkan potensi terbesar)
KLU_ORDER = [
    "INDUSTRI PENGOLAHAN",
    "KARYAWAN/PEGAWAI (KHUSUS DJP)",
    "AKTIVITAS PROFESIONAL, ILMIAH DAN TEKNIS",
    "PERDAGANGAN BESAR DAN ECERAN; REPARASI DAN PERAWATAN MOBIL DAN SEPEDA MOTOR",
    "AKTIVITAS KEUANGAN DAN ASURANSI",
    "PERTAMBANGAN DAN PENGGALIAN",
    "PENGADAAN LISTRIK, GAS, UAP/AIR PANAS DAN UDARA DINGIN",
    "REAL ESTAT",
    "PERTANIAN, KEHUTANAN, DAN PERIKANAN",
    "PENGANGKUTAN DAN PERGUDANGAN",
    "INFORMASI DAN KOMUNIKASI",
    "KONSTRUKSI",
    "AKTIVITAS KESEHATAN MANUSIA DAN AKTIVITAS SOSIAL",
    "AKTIVITAS JASA LAINNYA",
    "PENYEDIAAN AKOMODASI DAN PENYEDIAAN MAKAN MINUM",
    "AKTIVITAS PENYEWAAN DAN SEWA GUNA USAHA TANPA HAK OPSI, "
    "KETENAGAKERJAAN, AGEN PERJALANAN DAN PENUNJANG USAHA LAINNYA",
    "KESENIAN, HIBURAN DAN REKREASI",
    "PENDIDIKAN",
    "PENGADAAN AIR, PENGELOLAAN SAMPAH DAN DAUR ULANG, "
    "PEMBUANGAN DAN PEMBERSIHAN LIMBAH DAN SAMPAH",
    "AKTIVITAS RUMAH TANGGA SEBAGAI PEMBERI KERJA; AKTIVITAS YANG "
    "MENGHASILKAN BARANG DAN JASA OLEH RUMAH TANGGA YANG DIGUNAKAN "
    "UNTUK MEMENUHI KEBUTUHAN SENDIRI",
    "ADMINISTRASI PEMERINTAHAN, PERTAHANAN DAN JAMINAN SOSIAL WAJIB",
]

KLU_COLORS = [
    "#d95f02","#1b9e77","#7570b3","#e7298a","#66a61e",
    "#e6ab02","#a6761d","#1a56db","#e41a1c","#377eb8",
    "#4daf4a","#984ea3","#ff7f00","#a65628","#f781bf",
    "#666666","#00838f","#bf360c","#558b2f","#4527a0","#607d8b",
]

JENIS_COLORS = [
    '#0ea5e9','#10b981','#f59e0b','#8b5cf6','#ef4444',
    '#06b6d4','#84cc16','#f97316',
]

PUB_COLORS = {'KPDJP': '#1a56db', 'KANWIL': '#0d9488'}
PUB_LABELS = {'KPDJP': 'Kantor Pusat (KPDJP)', 'KANWIL': 'Kanwil'}

# =============================================================
# HTML TEMPLATE GABUNGAN
# =============================================================

DASH_HTML = r'''<div id="lha_dash-ctrl">
      <div class="pnrb-wrap">
        <span class="ctrl-lbl">Penerbit:</span>
        <label class="pnrb-item"><input type="checkbox" id="lha_ckKP" checked onchange="onFilter()"><label for="ckKP">Kantor Pusat</label></label>
        <label class="pnrb-item"><input type="checkbox" id="lha_ckKW" checked onchange="onFilter()"><label for="ckKW">Kanwil</label></label>
      </div>
      <div class="divider"></div>
      <span class="ctrl-lbl">Tahun Pajak:</span>
      <select class="ysel" id="lha_ysel" onchange="onFilter()">{{YEAR_OPTS}}</select>
      <span class="yb" id="lha_ybadge">ALL</span>
    </div>
    <div id="lha_kpistrip"></div>
    <div id="lha_dash-body">
      <div class="ttb">
        <div><div class="ttl">Data LHA per Kanwil &amp; KPP</div><div class="tmt" id="lha_tmeta">Klik Kanwil/KPP untuk expand detail</div></div>
        <div class="bg">
          <button class="btn" onclick="expAll()">&#9660; Buka Semua</button>
          <button class="btn" onclick="colAll()">&#9650; Tutup Semua</button>
        </div>
      </div>
      <div id="lha_scrollarea">
        <div class="vsl-wrap"><div class="vsl-lbl">SCROLL</div>
          <input type="range" class="vsl" id="lha_vsl" min="0" max="10000" value="0" oninput="syncV(this.value)">
        </div>
        <div id="lha_tblbox">
          <div id="lha_tblscroll">
            <table>
              <thead>
                <tr class="h1">
                  <th class="nc" rowspan="3">Kanwil / KPP / Penerbit</th>
                  <th rowspan="3" style="min-width:44px">Jml LHA</th>
                  <th rowspan="3" style="min-width:90px">Potensi LHA</th>
                  <th class="cs" rowspan="3"></th><th colspan="4">SP2DK</th>
                  <th class="cs" rowspan="3"></th><th colspan="4">LHP2DK</th>
                  <th class="cs" rowspan="3"></th><th colspan="6">Status LHA</th>
                  <th class="cs" rowspan="3"></th><th colspan="5">Jenis LHA</th>
                </tr>
                <tr class="h2">
                  <th colspan="2" style="border-right:1px solid rgba(200,168,75,.3)">Terbit SP2DK</th><th colspan="2">Belum Terbit SP2DK</th>
                  <th colspan="2" style="border-right:1px solid rgba(200,168,75,.3)">Terbit LHP2DK</th><th colspan="2">Belum Terbit LHP2DK</th>
                  <th colspan="2" style="border-right:1px solid rgba(200,168,75,.3)">Selesai</th><th colspan="4">Open</th>
                  <th colspan="5"></th>
                </tr>
                <tr class="h3">
                  <th>Jml</th><th style="min-width:82px">Pot. Awal</th>
                  <th>Jml</th><th style="min-width:82px">Pot. LHA</th>
                  <th>Jml</th><th style="min-width:82px">Pot. Akhir</th>
                  <th>Jml</th><th style="min-width:82px">Pot. Awal</th>
                  <th>Jml</th><th style="min-width:82px">Nilai Real.</th>
                  <th>Jml</th><th style="min-width:82px">Pot. LHA</th><th style="min-width:82px">Pot. Awal</th><th style="min-width:82px">Pot. Akhir</th>
                  <th style="min-width:50px">WP Grp</th><th style="min-width:50px">Lainnya</th><th style="min-width:44px">Joint</th><th style="min-width:52px">Reguler</th><th style="min-width:48px">Gugus</th>
                </tr>
              </thead>
              <tbody id="lha_tbody"></tbody>
            </table>
          </div>
          <div class="hsl-bar">
            <label>&#8646; Geser Kanan/Kiri</label>
            <input type="range" class="hsl" id="lha_hsl" min="0" max="10000" value="0" oninput="syncH(this.value)">
          </div>
        </div>
      </div>
    </div>'''
PETA_HTML = r'''<div id="lha_peta-wrap">
      <div id="lha_peta-sbar">
        <div class="sc"><span class="lb">Total LHA</span><span class="vl" id="lha_s-tot">&mdash;</span></div>
        <div class="sc"><span class="lb">Total Potensi LHA</span><span class="vl" id="lha_s-pot">&mdash;</span></div>
        <div class="sc"><span class="lb">Total Pembayaran</span><span class="vl" id="lha_s-bay">&mdash;</span></div>
        <div class="sc"><span class="lb">SP2DK Terbit</span><span class="vl" id="lha_s-sp2dk">&mdash;</span></div>
        <div class="sc"><span class="lb">KPP Terdampak</span><span class="vl" id="lha_s-kpp">&mdash;</span></div>
        <div class="sc"><span class="lb">Tampilan</span><span class="vl" id="lha_s-mod">Point KPP</span></div>
      </div>
      <div id="lha_peta-crow">
        <div id="lha_sidebar">
          <div class="sb-body">
            <div class="sb-sec">
              <div class="sb-sec-title">&#x1F5FA; Tampilan Peta</div>
              <div class="tg2">
                <button class="tb2 on" id="lha_tb-pt" onclick="setMode('point')">&#x1F4CD; Point KPP</button>
                <button class="tb2" id="lha_tb-ar" onclick="setMode('area')">&#x1F30F; Area Provinsi</button>
              </div>
            </div>
            <div class="sb-sec">
              <div class="sb-sec-title">&#x1F3A8; Metrik Warna</div>
              <div class="tg2">
                <button class="tb2 on" id="lha_m-seb" onclick="setMetric('sebaran')">Sebaran LHA</button>
                <button class="tb2" id="lha_m-pot" onclick="setMetric('potensi')">Potensi LHA</button>
              </div>
            </div>
            <div class="sb-sec">
              <div class="sb-sec-title">&#x1F3E2; Penerbit LHA</div>
              <div class="chips-wrap">{{PUB_CHIPS}}</div>
              <div style="font-size:9px;color:#8fa3b8;margin-top:5px;font-style:italic;">Kosong = tampilkan semua penerbit</div>
            </div>
            <div class="sb-sec">
              <div class="sb-sec-title">&#x1F4CB; Jenis LHA</div>
              <div class="chips-wrap">{{JEN_CHIPS}}</div>
              <div style="font-size:9px;color:#8fa3b8;margin-top:5px;font-style:italic;">Kosong = tampilkan semua jenis</div>
            </div>
            <div class="sb-sec">
              <div class="klu-hdr">
                <div class="sb-sec-title" style="margin-bottom:0;">&#x1F50D; Filter KLU</div>
                <div style="display:flex;align-items:center;gap:6px;">
                  <span id="lha_klu-count">Semua</span>
                  <button id="lha_klu-clear" onclick="clearKLU()">&#x2715; Reset</button>
                </div>
              </div>
              <div class="klu-note">Klik KLU untuk filter. Multi-pilih. Kosong = semua.</div>
              <div id="lha_kchips">{{KLU_CHIPS}}</div>
            </div>
          </div>
        </div>
        <button id="lha_stog" onclick="toggleSidebar()" title="Sembunyikan panel">&#x276E;</button>
        <div id="lha_marea"><div id="lha_map"></div><div id="lha_leg"></div></div>
      </div>
</div>'''
CSS_CODE = r''':root{
  --bg:#eef1f5;--sr:#fff;--s2:#f5f7fa;--bd:#dde2ea;--bds:#b8c2d0;
  --t1:#111827;--t2:#4b5563;--tm:#9ca3af;--ac:#1d4ed8;--al:#eff6ff;
  --djp-dark:#0a2346;--djp-gold:#c8a84b;--djp-gold2:#e8c96a;
  --fn:'DM Sans',sans-serif;--mo:'DM Mono',monospace;
}
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;overflow:hidden;font-family:var(--fn);background:var(--bg);color:var(--t1);font-size:12px}
body{display:flex;flex-direction:column}
/* MASTER HEADER */
#lha_master-hdr{background:var(--djp-dark);border-bottom:3px solid var(--djp-gold);flex-shrink:0;z-index:400;box-shadow:0 2px 8px rgba(0,0,0,.3)}
#lha_hdr-top{padding:8px 18px;display:flex;align-items:center;gap:12px}
.hdr-badge{background:var(--djp-gold);color:var(--djp-dark);font-size:11px;font-weight:800;letter-spacing:.5px;padding:4px 12px;border-radius:20px}
.hdr-titles h1{font-size:14px;font-weight:700;color:var(--djp-gold2);letter-spacing:-.2px}
.hdr-titles p{font-size:10px;color:rgba(200,168,75,.65)}
#lha_tab-bar{display:flex;padding:0 8px;gap:2px;background:rgba(0,0,0,.2)}
.tab-btn{font-family:var(--fn);font-size:11px;font-weight:700;padding:8px 22px;border:none;cursor:pointer;border-radius:6px 6px 0 0;letter-spacing:.2px;transition:all .2s}
.tab-btn.active{background:var(--djp-gold);color:var(--djp-dark)}
.tab-btn:not(.active){background:transparent;color:rgba(200,168,75,.55)}
.tab-btn:not(.active):hover{background:rgba(200,168,75,.12);color:var(--djp-gold2)}
/* CONTENT PANELS */
#lha_content-area{flex:1;overflow:hidden;display:flex;flex-direction:column}
.tab-panel{flex:1;display:none;overflow:hidden}
.tab-panel.active{display:flex;flex-direction:column}
/* DASHBOARD STYLES */
#lha_dash-ctrl{background:var(--djp-dark);padding:6px 18px;display:flex;align-items:center;gap:14px;flex-shrink:0}
.bx{display:flex;align-items:center;gap:10px}
.ctrl-lbl{font-size:11px;font-weight:600;color:var(--djp-gold);white-space:nowrap}
.ysel{font-family:var(--fn);font-size:12px;font-weight:500;padding:5px 9px;border:1.5px solid var(--djp-gold);border-radius:5px;background:var(--djp-dark);color:var(--djp-gold2);cursor:pointer;outline:none}
.yb{font-size:10px;font-weight:700;padding:3px 9px;border-radius:20px;background:rgba(200,168,75,.18);color:var(--djp-gold2)}
.pnrb-wrap{display:flex;align-items:center;gap:10px}
.pnrb-item{display:flex;align-items:center;gap:5px;cursor:pointer}
.pnrb-item input[type=checkbox]{width:14px;height:14px;accent-color:var(--djp-gold);cursor:pointer}
.pnrb-item label{font-size:11px;font-weight:600;color:var(--djp-gold2);cursor:pointer;white-space:nowrap}
.divider{width:1px;height:20px;background:rgba(200,168,75,.3)}
#lha_kpistrip{background:linear-gradient(135deg,#0d2d5a 0%,#0a2346 100%);border-bottom:2px solid var(--djp-gold);padding:6px 18px;display:flex;gap:7px;flex-shrink:0}
.kc{background:rgba(255,255,255,.06);border:1px solid rgba(200,168,75,.25);border-radius:5px;padding:5px 12px;flex:1;border-top:2px solid var(--djp-gold);min-width:0}
.kc.gn{border-top-color:#22c55e}.kc.am{border-top-color:#f59e0b}.kc.pu{border-top-color:#a78bfa}.kc.rd{border-top-color:#f87171}
.kl{font-size:8px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:rgba(200,168,75,.7);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.kv{font-size:13px;font-weight:700;font-family:var(--mo);color:#fff;letter-spacing:-.5px;white-space:nowrap}
.kc.gn .kv{color:#86efac}.kc.am .kv{color:#fcd34d}.kc.pu .kv{color:#c4b5fd}.kc.rd .kv{color:#fca5a5}
.ks2{font-size:8px;color:rgba(200,168,75,.5);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#lha_dash-body{flex:1;display:flex;flex-direction:column;overflow:hidden;padding:10px 16px 0}
.ttb{display:flex;align-items:center;justify-content:space-between;margin-bottom:7px;flex-shrink:0}
.ttl{font-size:12px;font-weight:700;color:var(--djp-dark)}.tmt{font-size:10px;color:var(--tm);margin-top:1px}
.bg{display:flex;gap:5px}
.btn{font-family:var(--fn);font-size:10px;font-weight:600;padding:4px 10px;border-radius:5px;border:1.5px solid var(--bds);background:var(--sr);color:var(--t2);cursor:pointer;transition:all .15s}
.btn:hover{background:var(--al);border-color:var(--ac);color:var(--ac)}
#lha_scrollarea{flex:1;display:flex;gap:6px;overflow:hidden;min-height:0;padding-bottom:10px}
.vsl-wrap{display:flex;flex-direction:column;align-items:center;width:16px;flex-shrink:0}
.vsl-lbl{font-size:7.5px;color:var(--tm);margin-bottom:3px;writing-mode:vertical-rl;user-select:none;letter-spacing:.5px}
input.vsl{writing-mode:vertical-lr;direction:rtl;flex:1;width:14px;accent-color:var(--djp-dark);cursor:pointer}
#lha_tblbox{flex:1;display:flex;flex-direction:column;border:1px solid var(--bd);border-radius:8px;background:var(--sr);overflow:hidden;min-width:0;box-shadow:0 2px 8px rgba(0,0,0,.06)}
#lha_tblscroll{flex:1;overflow:auto;min-height:0}
#lha_tblscroll::-webkit-scrollbar{width:5px;height:5px}
#lha_tblscroll::-webkit-scrollbar-track{background:var(--s2)}
#lha_tblscroll::-webkit-scrollbar-thumb{background:var(--bds);border-radius:3px}
.hsl-bar{padding:4px 8px;background:var(--s2);border-top:1px solid var(--bd);display:flex;align-items:center;gap:6px;flex-shrink:0}
.hsl-bar label{font-size:9px;color:var(--tm);user-select:none}
input.hsl{flex:1;accent-color:var(--djp-dark);cursor:pointer;height:4px}
table{border-collapse:collapse;white-space:nowrap;width:100%}
thead{position:sticky;top:0;z-index:50}
tr.h1 th{background:var(--djp-dark);color:var(--djp-gold);font-size:8.5px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;padding:7px 8px;text-align:center;border-right:1px solid rgba(200,168,75,.25)}
tr.h1 th.nc{text-align:left;border-right:2px solid rgba(200,168,75,.4)}
tr.h2 th{background:#122d5c;color:var(--djp-gold);font-size:8px;font-weight:600;padding:5px 8px;text-align:center;border-right:1px solid rgba(200,168,75,.2)}
tr.h3 th{background:#163570;color:rgba(200,168,75,.75);font-size:7.5px;font-weight:500;padding:3px 8px;text-align:center;border-right:1px solid rgba(200,168,75,.15);border-bottom:2px solid var(--djp-gold)}
thead th:last-child{border-right:none}
.nc{position:sticky;left:0;z-index:10;min-width:250px;max-width:250px;border-right:2px solid var(--bds)!important}
thead .nc{z-index:60;background:var(--djp-dark)!important;border-right:2px solid rgba(200,168,75,.4)!important}
.cs{background:var(--djp-dark)!important;width:3px;padding:0!important;border:none!important}
tbody td.cs{background:#dde5f0!important;width:3px;padding:0!important;border:none!important}
tbody td{padding:0 7px;height:26px;border-bottom:1px solid var(--bd);border-right:1px solid var(--bd);vertical-align:middle}
tbody td:last-child{border-right:none}
.num{font-family:var(--mo);font-size:11px;text-align:right}.nb{font-weight:600}
.nc-body{white-space:normal!important;word-break:break-word;line-height:1.3;padding-top:3px;padding-bottom:3px}
.rk{cursor:pointer}.rk td{background:#eef3ff;border-bottom:1px solid #c0ceec;min-height:32px}.rk td.nc{font-weight:700;font-size:11px;background:#eef3ff;min-height:32px}
.rk:hover td{background:var(--al)!important}.rk:hover td.nc{background:var(--al)!important}
.rko td{background:#f6f7fa}.rko td.nc{background:#f6f7fa}
.rks td{border-top:2px solid #a8bedc!important}
.rp td{background:#fbfcfe;height:24px}.rp td.nc{background:#fbfcfe;padding-left:14px;font-size:11px;color:var(--t2)}
.rpt td{background:#edf2fb;font-weight:700}.rpt td.nc{background:#edf2fb;color:var(--t1);font-size:11px}
.rkp{cursor:pointer}.rkp td{height:27px;background:#f0f5ff}.rkp td.nc{background:#f0f5ff;padding-left:20px;font-weight:600;font-size:11px;color:#334155}
.rkp:hover td{background:#e0eaff!important}.rkp:hover td.nc{background:#e0eaff!important}
.rkpp td{height:23px;background:#f7faff}.rkpp td.nc{background:#f7faff;padding-left:38px;font-size:10.5px;color:var(--tm)}
.rkpt td{height:24px;background:#ebf0fa;font-weight:700}.rkpt td.nc{background:#ebf0fa;padding-left:38px;font-size:10.5px;color:var(--t2)}
.tg{display:inline-flex;align-items:center;justify-content:center;width:13px;height:13px;border-radius:3px;background:var(--djp-dark);color:var(--djp-gold);font-size:7px;font-weight:900;margin-right:5px;flex-shrink:0;transition:transform .15s;vertical-align:middle}
.tg.op{transform:rotate(90deg)}.tg.sm{background:#64748b;color:#fff;width:11px;height:11px;font-size:6px}
.nw{display:flex;align-items:center}.kt{font-family:var(--mo);font-size:9px;color:var(--tm);margin-right:5px;min-width:26px;flex-shrink:0}
.bdg{font-size:8px;font-family:var(--mo);padding:1px 4px;border-radius:3px;margin-right:5px;flex-shrink:0}
.rz td{opacity:.38}.hd{display:none!important}
/* PETA STYLES */
#lha_peta-wrap{display:flex;flex-direction:column;height:100%;overflow:hidden;font-family:'Plus Jakarta Sans',sans-serif;background:#eef2f7}
#lha_peta-sbar{background:#f6f9fc;border-bottom:1px solid #dae1ea;padding:4px 18px;display:flex;overflow-x:auto;flex-shrink:0}
.sc{padding:4px 16px;border-right:1px solid #dae1ea;display:flex;flex-direction:column;gap:1px;min-width:110px;flex-shrink:0}
.sc:last-child{border-right:none}
.sc .lb{font-size:9px;font-weight:700;color:#8fa3b8;text-transform:uppercase;letter-spacing:.6px}
.sc .vl{font-size:14px;font-weight:800;color:#1a56db}
#lha_peta-crow{display:flex;flex:1;overflow:hidden;position:relative}
#lha_sidebar{width:280px;min-width:280px;background:#fff;border-right:1px solid #dae1ea;display:flex;flex-direction:column;overflow:hidden;transition:width .3s ease,min-width .3s ease;flex-shrink:0;box-shadow:2px 0 8px rgba(0,0,0,.04);z-index:10}
#lha_stog{position:absolute;left:280px;top:50%;transform:translateY(-50%);z-index:1001;width:20px;height:48px;background:#fff;border:1px solid #dae1ea;border-left:none;border-radius:0 6px 6px 0;cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:12px;color:#8fa3b8;transition:left .3s ease,background .15s;box-shadow:2px 0 6px rgba(0,0,0,.06);font-family:inherit}
#lha_stog:hover{background:#e8f0fe;color:#1a56db}
.sb-sec{border-bottom:1px solid #dae1ea;padding:10px 12px;flex-shrink:0}
.sb-sec-title{font-size:10px;font-weight:800;color:#8fa3b8;text-transform:uppercase;letter-spacing:.7px;margin-bottom:7px;display:flex;align-items:center;gap:5px}
.sb-body{flex:1;overflow-y:auto}
.tg2{display:flex;background:#eef2f7;border-radius:8px;border:1px solid #c2cedb;overflow:hidden;width:100%}
.tb2{flex:1;padding:6px 8px;font-size:11px;font-weight:700;cursor:pointer;border:none;background:transparent;color:#8fa3b8;font-family:inherit;transition:all .15s;white-space:nowrap;text-align:center}
.tb2.on{background:#1a56db;color:#fff}
.tb2:hover:not(.on){background:#f0f4fa;color:#1a56db}
.chips-wrap{display:flex;flex-wrap:wrap;gap:4px}
.chip{font-size:10px;font-weight:700;padding:4px 9px;border-radius:7px;cursor:pointer;border:2px solid transparent;transition:all .15s;color:#fff;opacity:.35;user-select:none}
.chip.on{opacity:1;box-shadow:0 2px 7px rgba(0,0,0,.2)}
.chip:hover{opacity:.7}.chip.on:hover{opacity:1}
#lha_kchips{display:flex;flex-direction:column;gap:4px}
.chip-klu{text-align:left;white-space:normal;line-height:1.3;padding:5px 10px;border-radius:8px}
.klu-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:7px}
#lha_klu-count{font-size:10px;font-weight:700;color:#1a56db}
#lha_klu-clear{display:none;align-items:center;gap:3px;font-size:10px;font-weight:700;color:#dc2626;cursor:pointer;background:#fef2f2;border:1px solid #fecaca;border-radius:5px;padding:2px 7px;white-space:nowrap;font-family:inherit}
#lha_klu-clear:hover{background:#fee2e2}
.klu-note{font-size:9.5px;color:#8fa3b8;margin-bottom:6px;line-height:1.5;font-style:italic}
#lha_marea{flex:1;position:relative;overflow:hidden}
#lha_map{width:100%;height:100%}
#lha_leg{display:none;position:absolute;bottom:20px;right:14px;z-index:1000;background:#fff;border:1px solid #dae1ea;border-radius:9px;padding:12px 14px;box-shadow:0 2px 10px rgba(0,0,0,.08);min-width:155px}
#lha_leg h4{font-size:10px;font-weight:800;color:#8fa3b8;text-transform:uppercase;letter-spacing:.6px;margin-bottom:7px}
.leaflet-popup-content-wrapper{background:#fff;border:1px solid #dae1ea;border-radius:11px;box-shadow:0 2px 10px rgba(0,0,0,.08);padding:0}
.leaflet-popup-content{margin:0!important;padding:0!important}
.leaflet-popup-tip{background:#fff}
.pb{padding:12px 15px;min-width:220px;max-width:315px}
.pb strong{display:block;font-size:12.5px;font-weight:800;color:#1c2b3a;margin-bottom:8px;padding-bottom:6px;border-bottom:2px solid #1a56db}
.pr{display:flex;justify-content:space-between;gap:8px;margin-top:4px;font-size:11px}
.lk{color:#8fa3b8;font-weight:500;white-space:nowrap}
.vk{color:#1a56db;font-weight:800}.vkh{color:#15803d;font-weight:800}
.vk2{color:#4a5e72;font-weight:600;font-size:10px;text-align:right;max-width:165px}
.pr-sep{border-top:1px dashed #dae1ea;margin:5px 0}
::-webkit-scrollbar{width:4px;height:4px}
::-webkit-scrollbar-thumb{background:#c2cedb;border-radius:4px}'''
JS_CODE = r'''// TAB SWITCHER
var _mapInit=false;
function switchTab(tab){
  document.querySelectorAll('.tab-panel').forEach(function(p){p.classList.remove('active');});
  document.querySelectorAll('.tab-btn').forEach(function(b){b.classList.remove('active');});
  document.getElementById('panel-'+tab).classList.add('active');
  document.getElementById('tab-btn-'+tab).classList.add('active');
  if(tab==='peta'&&!_mapInit){_mapInit=true;setTimeout(function(){if(typeof initPeta==='function')initPeta();},80);}
  if(tab==='peta'&&_mapInit&&typeof map!=='undefined'){setTimeout(function(){map.invalidateSize();},100);}
}
// DASHBOARD JS
var _D=(function(){var s=
{{DASH_DATA_JS}}
;return JSON.parse(atob(s));})();
var _cy='ALL',_ek={},_ep={};
function fR(v){if(!v||isNaN(v)||v===0)return'\u2014';if(v>=1e12)return(v/1e12).toFixed(2)+'\u00a0T';if(v>=1e9)return(v/1e9).toFixed(2)+'\u00a0M';if(v>=1e6)return(v/1e6).toFixed(0)+'\u00a0Jt';return v.toLocaleString('id-ID')}
function fI(v){if(!v||v===0)return'\u2014';return v.toLocaleString('id-ID')}
function zeroAgg(d){var r={};for(var k in d)r[k]=0;return r;}
function getEff(e){var kp=document.getElementById('lha_ckKP').checked,kw=document.getElementById('lha_ckKW').checked;if(kp&&kw)return e.tt;if(kp)return e.kp;if(kw)return e.kw;return zeroAgg(e.tt);}
function getNatEff(n){var kp=document.getElementById('lha_ckKP').checked,kw=document.getElementById('lha_ckKW').checked;if(kp&&kw)return n.all;if(kp)return n.kp;if(kw)return n.kw;return zeroAgg(n.all);}
function rKpi(){var n=getNatEff(_D.nat[_cy]||_D.nat['ALL']);var cards=[{l:'Total LHA',v:fI(n.n),s:'Seluruh nasional',c:''},{l:'Potensi LHA',v:fR(n.pot),s:'Nilai potensi',c:''},{l:'Terbit SP2DK',v:fI(n.sp_t),s:'Blm: '+fI(n.sp_b),c:'gn'},{l:'Pot. Awal SP2DK',v:fR(n.sp_t_v),s:'Terbit',c:'gn'},{l:'Terbit LHP2DK',v:fI(n.lhp_t),s:'Blm: '+fI(n.lhp_b),c:'pu'},{l:'Pot. LHP2DK',v:fR(n.lhp_t_v),s:'Terbit',c:'pu'},{l:'Selesai',v:fI(n.sel),s:'Real+TA+PPS+Close',c:'am'},{l:'Nilai Realisasi',v:fR(n.sel_rv),s:'Pembayaran',c:'am'},{l:'Open',v:fI(n.opn),s:'Pot: '+fR(n.opn_pot),c:'rd'}];document.getElementById('lha_kpistrip').innerHTML=cards.map(function(c){return'<div class="kc '+c.c+'"><div class="kl">'+c.l+'</div><div class="kv">'+c.v+'</div><div class="ks2">'+c.s+'</div></div>';}).join('');}
function cells(d){return'<td class="num nb">'+fI(d.n)+'</td><td class="num nb">'+fR(d.pot)+'</td><td class="cs"></td><td class="num">'+fI(d.sp_t)+'</td><td class="num">'+fR(d.sp_t_v)+'</td><td class="num">'+fI(d.sp_b)+'</td><td class="num">'+fR(d.sp_b_v)+'</td><td class="cs"></td><td class="num">'+fI(d.lhp_t)+'</td><td class="num">'+fR(d.lhp_t_v)+'</td><td class="num">'+fI(d.lhp_b)+'</td><td class="num">'+fR(d.lhp_b_v)+'</td><td class="cs"></td><td class="num">'+fI(d.sel)+'</td><td class="num">'+fR(d.sel_rv)+'</td><td class="num">'+fI(d.opn)+'</td><td class="num">'+fR(d.opn_pot)+'</td><td class="num">'+fR(d.opn_pa)+'</td><td class="num">'+fR(d.opn_pak)+'</td><td class="cs"></td><td class="num">'+fI(d.wpg)+'</td><td class="num">'+fI(d.lnx)+'</td><td class="num">'+fI(d.jnt)+'</td><td class="num">'+fI(d.reg)+'</td><td class="num">'+fI(d.gug)+'</td>';}
function nc(inner){return'<td class="nc nc-body">'+inner+'</td>';}
function rTable(yr){var yd=_D.dat[yr]||_D.dat['ALL'];var ckKP=document.getElementById('lha_ckKP').checked,ckKW=document.getElementById('lha_ckKW').checked;var rows=[];_D.kwl.forEach(function(kw,idx){var kwl=String(kw.kwl),kwd=yd[kwl];if(!kwd)return;var isE=!!_ek[kwl],eff=getEff(kwd);var sn=kwd.nm.replace(/^Kantor Wilayah /i,'').replace(/^Kanwil /i,'');rows.push('<tr class="rk'+(idx%2===1?' rko':'')+(idx>0?' rks':'')+'" data-kwl="'+kwl+'">'+nc('<div class="nw"><span class="tg'+(isE?' op':'')+'" id="ic'+kwl+'">\u25ba</span><span class="kt">'+String(kw.kwl).padStart(3,'0')+'</span>'+sn+'</div>')+cells(eff)+'</tr>');var hc=isE?'':'hd';if(ckKP&&ckKW){rows.push('<tr class="rp kwc-'+kwl+' '+hc+'">'+nc('<div class="nw" style="padding-left:12px"><span class="bdg" style="background:#eff6ff;color:#1d4ed8">KPDJP</span>Kantor Pusat</div>')+cells(kwd.kp)+'</tr>');rows.push('<tr class="rp kwc-'+kwl+' '+hc+'">'+nc('<div class="nw" style="padding-left:12px"><span class="bdg" style="background:#ecfdf5;color:#059669">KANWIL</span>Kantor Wilayah</div>')+cells(kwd.kw)+'</tr>');rows.push('<tr class="rp rpt kwc-'+kwl+' '+hc+'">'+nc('<div class="nw" style="padding-left:12px"><span class="bdg" style="background:#f0f4fb;color:#475569">TOTAL</span>Total Kanwil</div>')+cells(kwd.tt)+'</tr>');}kwd.pp.forEach(function(kp){var kid=kwl+'_'+String(kp.c),isKE=!!_ep[kid];var kH=isE?'':'hd',kcH=(isE&&isKE)?'':'hd';var effKp=getEff(kp),zero=effKp.n===0;rows.push('<tr class="rkp kwc-'+kwl+' '+(zero?'rz ':'')+kH+'" data-kid="'+kid+'">'+nc('<div class="nw" style="padding-left:18px"><span class="tg sm'+(isKE?' op':'')+'" id="ick'+kid+'">\u25ba</span><span class="kt">'+String(kp.c).padStart(3,'0')+'</span>'+kp.nm+'</div>')+cells(effKp)+'</tr>');if(ckKP&&ckKW){rows.push('<tr class="rkpp kwc-'+kwl+' kpc-'+kid+' '+kcH+'">'+nc('<div class="nw" style="padding-left:36px"><span class="bdg" style="background:#eff6ff;color:#1d4ed8">KPDJP</span></div>')+cells(kp.kp)+'</tr>');rows.push('<tr class="rkpp kwc-'+kwl+' kpc-'+kid+' '+kcH+'">'+nc('<div class="nw" style="padding-left:36px"><span class="bdg" style="background:#ecfdf5;color:#059669">KANWIL</span></div>')+cells(kp.kw)+'</tr>');rows.push('<tr class="rkpt kwc-'+kwl+' kpc-'+kid+' '+kcH+'">'+nc('<div class="nw" style="padding-left:36px;font-weight:700;font-size:10.5px;color:var(--t2)">Total KPP</div>')+cells(kp.tt)+'</tr>');}});});document.getElementById('lha_tbody').innerHTML=rows.join('');var pLabel=(ckKP&&ckKW)?'Semua Penerbit':(ckKP?'Kantor Pusat':(ckKW?'Kanwil':'\u2014'));document.getElementById('lha_tmeta').textContent=(_cy==='ALL'?'Semua Tahun Pajak':'TP '+_cy)+' \u00b7 '+pLabel+' \u00b7 Klik untuk expand';iSlider();}
document.addEventListener('click',function(e){var kr=e.target.closest('[data-kwl]'),kp=e.target.closest('[data-kid]');if(kp){tKpp(kp.getAttribute('data-kid'));return;}if(kr){tKwl(kr.getAttribute('data-kwl'));return;}});
function tKwl(kwl){_ek[kwl]=!_ek[kwl];var show=_ek[kwl];document.querySelectorAll('.kwc-'+kwl).forEach(function(el){var kid=[...el.classList].map(function(c){return c.startsWith('kpc-')?c.replace('kpc-',''):null;}).find(Boolean);if(kid){if(show&&_ep[kid])el.classList.remove('hd');else el.classList.add('hd');}else{show?el.classList.remove('hd'):el.classList.add('hd');}});var ic=document.getElementById('ic'+kwl);if(ic)ic.classList.toggle('op',show);}
function tKpp(kid){_ep[kid]=!_ep[kid];var show=_ep[kid];document.querySelectorAll('.kpc-'+kid).forEach(function(el){show?el.classList.remove('hd'):el.classList.add('hd');});var ic=document.getElementById('ick'+kid);if(ic)ic.classList.toggle('op',show);}
function expAll(){_D.kwl.forEach(function(kw){_ek[String(kw.kwl)]=true;});rTable(_cy);}
function colAll(){_ek={};_ep={};rTable(_cy);}
function onFilter(){var kp=document.getElementById('lha_ckKP'),kw=document.getElementById('lha_ckKW');if(!kp.checked&&!kw.checked){kp.checked=true;return;}_cy=document.getElementById('lha_ysel').value;document.getElementById('lha_ybadge').textContent=_cy;_ek={};_ep={};rKpi();rTable(_cy);}
function iSlider(){var ts=document.getElementById('lha_tblscroll'),h=document.getElementById('lha_hsl'),v=document.getElementById('lha_vsl');setTimeout(function(){h.max=Math.max(1,ts.scrollWidth-ts.clientWidth);v.max=Math.max(1,ts.scrollHeight-ts.clientHeight);h.value=ts.scrollLeft;v.value=ts.scrollTop;},50);}
function syncH(val){document.getElementById('lha_tblscroll').scrollLeft=parseInt(val);}
function syncV(val){document.getElementById('lha_tblscroll').scrollTop=parseInt(val);}
document.getElementById('lha_tblscroll').addEventListener('scroll',function(){var h=document.getElementById('lha_hsl'),v=document.getElementById('lha_vsl');h.max=Math.max(1,this.scrollWidth-this.clientWidth);v.max=Math.max(1,this.scrollHeight-this.clientHeight);h.value=this.scrollLeft;v.value=this.scrollTop;});
rKpi();rTable('ALL');setTimeout(iSlider,400);
// PETA JS
window._PL="{{PETA_PAYLOAD_B64}}";
window._GJ="{{GEO_B64}}";
window._KR="{{KREF_B64}}";
window._KWP="{{KWP_B64}}";
function initPeta(){(function(){var s=document.createElement('script');s.textContent=atob("{{JS_B64}}");document.head.appendChild(s);})();}'''
JS_TEMPLATE = r'''
const _DPAYLOAD = (function(){{ return JSON.parse(atob("{{DASH_PAYLOAD}}")); }})();
const _DATA = (function(){{
    var s = "{{PETA_B64}}";
    if (!s) return [];
    var p = "";
    var chunks = s.split(",");
    for(var i=0; i<chunks.length; i++){{ p += chunks[i]; }}
    return JSON.parse(decodeURIComponent(escape(atob(p))));
}})();
'''
def build_agg(data):
    kat_sel = ['Realisasi', 'TA', 'PPS', 'Close Tanpa Realisasi']
    kat_opn = ['Belum Realisasi', 'Belum Tindak Lanjut', 'Usul Bukper', 'Usul Riksus']
    m_sp_t  = data['Status SP2DK']  == 'Terbit SP2DK'
    m_sp_b  = data['Status SP2DK']  == 'Belum Terbit SP2DK'
    m_lhp_t = data['Status LHP2DK'] == 'Terbit LHP2DK'
    m_lhp_b = data['Status LHP2DK'] == 'Belum Terbit LHP2DK'
    m_sel   = data['KATEGORI'].isin(kat_sel)
    m_opn   = data['KATEGORI'].isin(kat_opn)
    return {
        'n':       int(len(data)),
        'pot':     float(data['POTENSI_LHA'].sum()),
        'sp_t':    int(m_sp_t.sum()),
        'sp_t_v':  float(data.loc[m_sp_t,  'POTENSI_AWAL'].sum()),
        'sp_b':    int(m_sp_b.sum()),
        'sp_b_v':  float(data.loc[m_sp_b,  'POTENSI_LHA'].sum()),
        'lhp_t':   int(m_lhp_t.sum()),
        'lhp_t_v': float(data.loc[m_lhp_t, 'POTENSI_AKHIR'].sum()),
        'lhp_b':   int(m_lhp_b.sum()),
        'lhp_b_v': float(data.loc[m_lhp_b, 'POTENSI_AWAL'].sum()),
        'sel':     int(m_sel.sum()),
        'sel_rv':  float(data['NILAI_PEMBAYARAN'].sum()),
        'opn':     int(m_opn.sum()),
        'opn_pot': float(data.loc[m_opn, 'POTENSI_LHA'].sum()),
        'opn_pa':  float(data.loc[m_opn, 'POTENSI_AWAL'].sum()),
        'opn_pak': float(data.loc[m_opn, 'POTENSI_AKHIR'].sum()),
        'wpg':  int((data['JENIS_LHA'] == 'WP GROUP').sum()),
        'lnx':  int((data['JENIS_LHA'] == 'Lainnya').sum()),
        'jnt':  int((data['JENIS_LHA'] == 'Joint Analisis').sum()),
        'reg':  int((data['JENIS_LHA'] == 'Reguler').sum()),
        'gug':  int((data['JENIS_LHA'] == 'Gugus Tugas').sum()),
        'd_real': int((data['KATEGORI'] == 'Realisasi').sum()),
        'd_ta':   int((data['KATEGORI'] == 'TA').sum()),
        'd_pps':  int((data['KATEGORI'] == 'PPS').sum()),
        'd_cls':  int((data['KATEGORI'] == 'Close Tanpa Realisasi').sum()),
        'd_blr':  int((data['KATEGORI'] == 'Belum Realisasi').sum()),
        'd_blt':  int((data['KATEGORI'] == 'Belum Tindak Lanjut').sum()),
        'd_ubk':  int((data['KATEGORI'] == 'Usul Bukper').sum()),
        'd_urk':  int((data['KATEGORI'] == 'Usul Riksus').sum()),
    }


def dash_process_data(df, df_map, col_map, progress_cb=None):
    """Build the full JSON payload from dataframes.
    Note: df must already have standard column names applied before calling."""
    # Ensure numeric types
    for c in ['POTENSI_LHA', 'POTENSI_AWAL', 'POTENSI_AKHIR', 'NILAI_PEMBAYARAN']:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
    df['KPP']         = pd.to_numeric(df['KPP'], errors='coerce').fillna(0).astype(int)
    df['KWL']         = pd.to_numeric(df['KWL'], errors='coerce').fillna(0).astype(int)
    df['TAHUN_PAJAK'] = pd.to_numeric(df['TAHUN_PAJAK'], errors='coerce')

    years = ['ALL'] + sorted([int(y) for y in df['TAHUN_PAJAK'].dropna().unique()])
    result = {}
    national = {}
    total_steps = len(years)

    for i, yr in enumerate(years):
        if progress_cb:
            progress_cb(int(i / total_steps * 85), f"Memproses tahun {yr}...")
        d_yr = df if yr == 'ALL' else df[df['TAHUN_PAJAK'] == yr]
        result[str(yr)] = {}
        national[str(yr)] = {
            'all': build_agg(d_yr),
            'kp':  build_agg(d_yr[d_yr['PENERBIT'] == 'KPDJP']),
            'kw':  build_agg(d_yr[d_yr['PENERBIT'] == 'KANWIL']),
        }
        for kwl_code, kwl_grp in df_map.groupby('kwl'):
            kwl = int(kwl_code)
            kwl_name = kwl_grp.iloc[0]['kwl_name']
            kpp_list = kwl_grp[['kpp', 'kpp_name']].sort_values('kpp')
            dk = d_yr[d_yr['KWL'] == kwl]
            entry = {
                'kwl': kwl, 'nm': kwl_name,
                'kp': build_agg(dk[dk['PENERBIT'] == 'KPDJP']),
                'kw': build_agg(dk[dk['PENERBIT'] == 'KANWIL']),
                'tt': build_agg(dk),
                'pp': []
            }
            for _, kr in kpp_list.iterrows():
                kc = int(kr['kpp'])
                dk2 = dk[dk['KPP'] == kc]
                entry['pp'].append({
                    'c': kc, 'nm': kr['kpp_name'],
                    'kp': build_agg(dk2[dk2['PENERBIT'] == 'KPDJP']),
                    'kw': build_agg(dk2[dk2['PENERBIT'] == 'KANWIL']),
                    'tt': build_agg(dk2),
                })
            result[str(yr)][str(kwl)] = entry

    kwl_list = df_map.groupby('kwl').first().reset_index()[['kwl', 'kwl_name']].sort_values('kwl').to_dict('records')
    return {'nat': national, 'kwl': kwl_list, 'dat': result}, years


def validate_columns(df, log):
    """
    Cari setiap kolom wajib berdasarkan nama (case-insensitive).
    Return: (col_map dict, missing list)
    col_map = {REQUIRED_NAME: actual_col_name_in_df}
    """
    cols_upper = {c.strip().upper(): c for c in df.columns}
    mapping, missing = {}, []
    for req in PETA_REQUIRED_COLS:
        if req in cols_upper:
            mapping[req] = cols_upper[req]
        else:
            missing.append(req)
    if missing:
        log(f"   ⚠️  Kolom tidak ditemukan : {', '.join(missing)}")
    else:
        log(f"   ✅ Semua {len(PETA_REQUIRED_COLS)} kolom wajib ditemukan")
    return mapping, missing


def peta_process_data(df, col_map, log):
    """
    Bersihkan dan agregasi data LHA menjadi payload JSON terkompresi.
    Return: (payload_b64, pub_vals, jenis_vals, klu_vals)
    """
    log("📊 Memproses data...")

    def col(name):
        return col_map.get(name, name)

    # Standardize key columns (upper, strip, fix KPP names)
    df['_KPP']   = (df[col('NAMA_KPP')].fillna('')
                    .str.strip().str.upper().replace(KPP_NAME_FIX))
    df['_KWL']   = df[col('KANWIL')].fillna('').str.strip().str.upper()
    df['_PUB']   = df[col('PENERBIT')].fillna('').str.strip().str.upper()
    df['_JENIS'] = df[col('JENIS_LHA')].fillna('').str.strip().str.upper()
    df['_KLU']   = df[col('NM_KATEGORI')].fillna('').str.strip().str.upper()

    log(f"   Total baris  : {len(df):,}")
    log(f"   KPP unik     : {df['_KPP'].nunique()}")
    log(f"   Kanwil unik  : {df['_KWL'].nunique()}")

    # Granular aggregation: (KPP × Kanwil × Penerbit × Jenis × KLU)
    log("📐 Agregasi per KPP × Penerbit × Jenis × KLU...")
    grp = df.groupby(
        ['_KPP', '_KWL', '_PUB', '_JENIS', '_KLU'], dropna=False
    )

    rows = []
    for (kpp, kwl, pub, jenis, klu), g in grp:
        def s(c):
            return int(g[col(c)].sum()) if g[col(c)].notna().any() else 0
        rows.append({
            'k': str(kpp),   'w': str(kwl),
            'p': str(pub),   'j': str(jenis),
            'l': str(klu),   'c': int(len(g)),
            'pt': s('POTENSI_LHA'),
            'pa': s('POTENSI_AWAL'),
            'pk': s('POTENSI_AKHIR'),
            'by': s('NILAI_PEMBAYARAN'),
            'sp': int(g[col('NO_SP2DK')].notna().sum()),
            'lh': int(g[col('NO_LHP2DK')].notna().sum()),
        })

    log(f"   Granular rows: {len(rows):,}")

    # Compress: replace repeated strings with integer indices
    log("🗜️  Mengkompresi payload...")
    kpp_l   = sorted(set(r['k'] for r in rows))
    kwl_l   = sorted(set(r['w'] for r in rows))
    pub_l   = sorted(set(r['p'] for r in rows))
    jenis_l = sorted(set(r['j'] for r in rows))
    klu_l   = sorted(set(r['l'] for r in rows))

    ki = {v: i for i, v in enumerate(kpp_l)}
    wi = {v: i for i, v in enumerate(kwl_l)}
    pi = {v: i for i, v in enumerate(pub_l)}
    ji = {v: i for i, v in enumerate(jenis_l)}
    li = {v: i for i, v in enumerate(klu_l)}

    compact = [
        [ki[r['k']], wi[r['w']], pi[r['p']], ji[r['j']], li[r['l']],
         r['c'], r['pt'], r['pa'], r['pk'], r['by'], r['sp'], r['lh']]
        for r in rows
    ]

    payload     = {'kpp': kpp_l, 'kwl': kwl_l, 'pub': pub_l,
                   'jenis': jenis_l, 'klu': klu_l, 'rows': compact}
    pj          = json.dumps(payload, ensure_ascii=False, separators=(',', ':'))
    payload_b64 = base64.b64encode(pj.encode()).decode()

    log(f"   Payload      : {len(pj)/1024:.0f} KB → base64: {len(payload_b64)/1024:.0f} KB")
    log(f"   Penerbit     : {pub_l}")
    log(f"   Jenis LHA    : {jenis_l}")
    log(f"   KLU kategori : {len(klu_l)}")

    total_pot = sum(r[6] for r in compact)
    log(f"   Total Potensi: Rp {total_pot/1e12:.2f} T")

    return payload_b64, pub_l, jenis_l, klu_l


# =============================================================
# HTML GENERATION
# =============================================================

def _chip(item, color, chip_class, fn, label=None):
    label = label or item
    short = (label[:34] + '…') if len(label) > 35 else label
    esc   = item.replace("'", "\\'").replace('"', '&quot;')
    return (f'<span class="chip {chip_class}" '
            f'style="background:{color};border-color:{color};" '
            f'title="{label}" onclick="{fn}(\'{esc}\',this)">'
            f'{short}</span>\n')


def generate_html(payload_b64, pub_vals, jenis_vals, klu_vals,
                  map_title, log):
    """Susun HTML lengkap identik dengan Peta LHA DJP versi terbaru."""
    log("🏗️  Menyusun HTML...")

    # Penerbit chips
    pub_chips = ''
    for pub in pub_vals:
        pub_chips += _chip(pub, PUB_COLORS.get(pub, '#555'),
                           'chip-pub', 'togglePub',
                           PUB_LABELS.get(pub, pub))

    # Jenis LHA chips
    jen_chips = ''
    for i, jen in enumerate(jenis_vals):
        jen_chips += _chip(jen, JENIS_COLORS[i % len(JENIS_COLORS)],
                           'chip-jen', 'toggleJenis')

    # KLU chips: urut sesuai KLU_ORDER, sisanya tambah di bawah
    klu_upper = {k.upper() for k in klu_vals}
    ordered   = [k for k in KLU_ORDER if k.upper() in klu_upper]
    extra     = [k for k in klu_vals
                 if k.upper() not in {x.upper() for x in KLU_ORDER}]
    klu_chips = ''
    for i, klu in enumerate(ordered + extra):
        klu_chips += _chip(klu, KLU_COLORS[i % len(KLU_COLORS)],
                           'chip-klu', 'toggleKLU')

    html = f'''<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>{map_title}</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{{--bg:#eef2f7;--surf:#fff;--surf2:#f6f9fc;--brd:#dae1ea;--brd2:#c2cedb;
--t1:#1c2b3a;--t2:#4a5e72;--t3:#8fa3b8;--acc:#1a56db;--acc2:#1040b0;
--acc-lt:#e8f0fe;--red:#dc2626;--grn:#15803d;
--sh:0 2px 10px rgba(0,0,0,.08),0 6px 24px rgba(0,0,0,.05);
--sh-s:0 1px 3px rgba(0,0,0,.07);--r:9px;}}
*{{margin:0;padding:0;box-sizing:border-box;}}
html,body{{height:100%;font-family:'Plus Jakarta Sans',sans-serif;
background:var(--bg);color:var(--t1);overflow:hidden;}}
#app{{display:grid;grid-template-rows:auto auto 1fr;height:100vh;}}
#crow{{display:flex;height:100%;overflow:hidden;position:relative;}}
#hdr{{background:var(--surf);border-bottom:3px solid var(--acc);
padding:8px 18px;display:flex;align-items:center;gap:10px;
flex-wrap:wrap;box-shadow:var(--sh-s);}}
.logo{{background:var(--acc);color:#fff;font-size:11px;font-weight:800;
letter-spacing:.8px;padding:4px 11px;border-radius:20px;flex-shrink:0;}}
.sec{{background:#fef2f2;color:var(--red);font-size:10px;font-weight:800;
padding:2px 8px;border-radius:4px;border:1px solid #fecaca;letter-spacing:.4px;}}
#sbar{{background:var(--surf2);border-bottom:1px solid var(--brd);
padding:4px 18px;display:flex;overflow-x:auto;flex-shrink:0;}}
.sc{{padding:4px 16px;border-right:1px solid var(--brd);
display:flex;flex-direction:column;gap:1px;min-width:110px;flex-shrink:0;}}
.sc:last-child{{border-right:none;}}
.sc .lb{{font-size:9px;font-weight:700;color:var(--t3);
text-transform:uppercase;letter-spacing:.6px;}}
.sc .vl{{font-size:14px;font-weight:800;color:var(--acc);}}
#sidebar{{width:280px;min-width:280px;background:var(--surf);
border-right:1px solid var(--brd);display:flex;flex-direction:column;
overflow:hidden;transition:width .3s ease,min-width .3s ease;
flex-shrink:0;box-shadow:2px 0 8px rgba(0,0,0,.04);z-index:10;}}
#stog{{position:absolute;left:280px;top:50%;transform:translateY(-50%);
z-index:1001;width:20px;height:48px;background:var(--surf);
border:1px solid var(--brd);border-left:none;border-radius:0 6px 6px 0;
cursor:pointer;display:flex;align-items:center;justify-content:center;
font-size:12px;color:var(--t3);transition:left .3s ease,background .15s;
box-shadow:2px 0 6px rgba(0,0,0,.06);font-family:inherit;}}
#stog:hover{{background:var(--acc-lt);color:var(--acc);}}
.sb-sec{{border-bottom:1px solid var(--brd);padding:10px 12px;flex-shrink:0;}}
.sb-sec-title{{font-size:10px;font-weight:800;color:var(--t3);
text-transform:uppercase;letter-spacing:.7px;margin-bottom:7px;
display:flex;align-items:center;gap:5px;}}
.sb-body{{flex:1;overflow-y:auto;}}
.tg{{display:flex;background:var(--bg);border-radius:8px;
border:1px solid var(--brd2);overflow:hidden;width:100%;}}
.tb2{{flex:1;padding:6px 8px;font-size:11px;font-weight:700;cursor:pointer;
border:none;background:transparent;color:var(--t3);font-family:inherit;
transition:all .15s;white-space:nowrap;text-align:center;}}
.tb2.on{{background:var(--acc);color:#fff;}}
.tb2:hover:not(.on){{background:#f0f4fa;color:var(--acc);}}
.chips-wrap{{display:flex;flex-wrap:wrap;gap:4px;}}
.chip{{font-size:10px;font-weight:700;padding:4px 9px;border-radius:7px;
cursor:pointer;border:2px solid transparent;transition:all .15s;
color:#fff;opacity:.35;user-select:none;}}
.chip.on{{opacity:1;box-shadow:0 2px 7px rgba(0,0,0,.2);}}
.chip:hover{{opacity:.7;}}
.chip.on:hover{{opacity:1;}}
#kchips{{display:flex;flex-direction:column;gap:4px;}}
.chip-klu{{text-align:left;white-space:normal;line-height:1.3;
padding:5px 10px;border-radius:8px;}}
.klu-hdr{{display:flex;align-items:center;justify-content:space-between;
margin-bottom:7px;}}
#klu-count{{font-size:10px;font-weight:700;color:var(--acc);}}
#klu-clear{{display:none;align-items:center;gap:3px;font-size:10px;
font-weight:700;color:var(--red);cursor:pointer;background:#fef2f2;
border:1px solid #fecaca;border-radius:5px;padding:2px 7px;
white-space:nowrap;font-family:inherit;}}
#klu-clear:hover{{background:#fee2e2;}}
.klu-note{{font-size:9.5px;color:var(--t3);margin-bottom:6px;
line-height:1.5;font-style:italic;}}
#marea{{flex:1;position:relative;overflow:hidden;}}
#map{{width:100%;height:100%;}}
#leg{{display:none;position:absolute;bottom:20px;right:14px;z-index:1000;
background:var(--surf);border:1px solid var(--brd);border-radius:var(--r);
padding:12px 14px;box-shadow:var(--sh);min-width:155px;}}
#leg h4{{font-size:10px;font-weight:800;color:var(--t3);
text-transform:uppercase;letter-spacing:.6px;margin-bottom:7px;}}
.leaflet-popup-content-wrapper{{background:var(--surf);border:1px solid var(--brd);
border-radius:11px;box-shadow:var(--sh);padding:0;}}
.leaflet-popup-content{{margin:0!important;padding:0!important;}}
.leaflet-popup-tip{{background:var(--surf);}}
.pb{{padding:12px 15px;min-width:220px;max-width:315px;}}
.pb strong{{display:block;font-size:12.5px;font-weight:800;color:var(--t1);
margin-bottom:8px;padding-bottom:6px;border-bottom:2px solid var(--acc);}}
.pr{{display:flex;justify-content:space-between;gap:8px;margin-top:4px;font-size:11px;}}
.lk{{color:var(--t3);font-weight:500;white-space:nowrap;}}
.vk{{color:var(--acc);font-weight:800;}}
.vkh{{color:var(--grn);font-weight:800;}}
.vk2{{color:var(--t2);font-weight:600;font-size:10px;text-align:right;max-width:165px;}}
.pr-sep{{border-top:1px dashed var(--brd);margin:5px 0;}}
::-webkit-scrollbar{{width:4px;height:4px;}}
::-webkit-scrollbar-thumb{{background:var(--brd2);border-radius:4px;}}
</style>
</head>
<body>
<div id="app">
<div id="hdr">
  <div class="logo">DJP</div>
  <span style="font-size:13.5px;font-weight:700;">{map_title}</span>
  <span class="sec">&#x1F512; RAHASIA</span>
</div>
<div id="sbar">
  <div class="sc"><span class="lb">Total LHA</span>
    <span class="vl" id="s-tot">&mdash;</span></div>
  <div class="sc"><span class="lb">Total Potensi LHA</span>
    <span class="vl" id="s-pot">&mdash;</span></div>
  <div class="sc"><span class="lb">Total Pembayaran</span>
    <span class="vl" id="s-bay">&mdash;</span></div>
  <div class="sc"><span class="lb">SP2DK Terbit</span>
    <span class="vl" id="s-sp2dk">&mdash;</span></div>
  <div class="sc"><span class="lb">KPP Terdampak</span>
    <span class="vl" id="s-kpp">&mdash;</span></div>
  <div class="sc"><span class="lb">Tampilan</span>
    <span class="vl" id="s-mod">Point KPP</span></div>
</div>
<div id="crow">
  <div id="sidebar">
    <div class="sb-body">
      <div class="sb-sec">
        <div class="sb-sec-title">&#x1F5FA; Tampilan Peta</div>
        <div class="tg">
          <button class="tb2 on" id="tb-pt" onclick="setMode('point')">
            &#x1F4CD; Point KPP</button>
          <button class="tb2" id="tb-ar" onclick="setMode('area')">
            &#x1F30F; Area Provinsi</button>
        </div>
      </div>
      <div class="sb-sec">
        <div class="sb-sec-title">&#x1F3A8; Metrik Warna</div>
        <div class="tg">
          <button class="tb2 on" id="m-seb" onclick="setMetric('sebaran')">
            Sebaran LHA</button>
          <button class="tb2" id="m-pot" onclick="setMetric('potensi')">
            Potensi LHA</button>
        </div>
      </div>
      <div class="sb-sec">
        <div class="sb-sec-title">&#x1F3E2; Penerbit LHA</div>
        <div class="chips-wrap">
{pub_chips}        </div>
        <div style="font-size:9px;color:var(--t3);margin-top:5px;font-style:italic;">
          Kosong = tampilkan semua penerbit</div>
      </div>
      <div class="sb-sec">
        <div class="sb-sec-title">&#x1F4CB; Jenis LHA</div>
        <div class="chips-wrap">
{jen_chips}        </div>
        <div style="font-size:9px;color:var(--t3);margin-top:5px;font-style:italic;">
          Kosong = tampilkan semua jenis</div>
      </div>
      <div class="sb-sec">
        <div class="klu-hdr">
          <div class="sb-sec-title" style="margin-bottom:0;">
            &#x1F50D; Filter KLU</div>
          <div style="display:flex;align-items:center;gap:6px;">
            <span id="klu-count">Semua</span>
            <button id="klu-clear" onclick="clearKLU()">&#x2715; Reset</button>
          </div>
        </div>
        <div class="klu-note">Klik KLU untuk filter. Multi-pilih. Kosong = semua.</div>
        <div id="kchips">
{klu_chips}        </div>
      </div>
    </div>
  </div>
  <button id="stog" onclick="toggleSidebar()" title="Sembunyikan panel">&#x276E;</button>
  <div id="marea">
    <div id="map"></div>
    <div id="leg"></div>
  </div>
</div>
</div>
<script>
window._PL  = "{payload_b64}";
window._GJ  = "{GEO_B64}";
window._KR  = "{KREF_B64}";
window._KWP = "{KWP_B64}";
(function(){{
  var s=document.createElement('script');
  s.textContent=atob("{JS_B64}");
  document.head.appendChild(s);
}})();
</script>
</body>
</html>'''

    log(f"✅ HTML selesai: {len(html)/1024:.0f} KB")
    return html

# =============================================================
# CHIP HELPER
# =============================================================
def _chip(item, color, chip_class, fn, label=None):
    label = label or item
    short = (label[:34] + "\u2026") if len(label) > 35 else label
    esc   = item.replace("'", "\\'").replace('"', '&quot;')
    return (f'<span class="chip {chip_class}" '
            f'style="background:{color};border-color:{color};" '
            f'title="{label}" onclick="{fn}(\'{esc}\',this)">'
            f'{short}</span>\n')

# =============================================================
# BUILD COMBINED HTML
# =============================================================



def build_lha_module(dash_payload, dash_years, peta_b64, pub_vals, jenis_vals, klu_vals, title, log):
    yopts = "<option value='ALL'>ALL (Semua Tahun)</option>"
    dash_years_clean = [int(y) for y in dash_years if str(y) != 'ALL']
    for y in sorted(dash_years_clean, reverse=True):
        yopts += f"<option value='{y}'>{y}</option>"

    dash_html_content = DASH_HTML.replace('{{YEAR_OPTS}}', yopts)

    def _build_opts(vals):
        return "".join([f"<option value='{v}'>{v}</option>" for v in vals])

    peta_html_content = PETA_HTML.replace('{{PUB_OPTS}}', _build_opts(pub_vals))
    peta_html_content = peta_html_content.replace('{{JENIS_OPTS}}', _build_opts(jenis_vals))
    peta_html_content = peta_html_content.replace('{{KLU_OPTS}}', _build_opts(klu_vals))

    import json
    import base64
    dash_payload_str = dash_payload if isinstance(dash_payload, str) else base64.b64encode(json.dumps(dash_payload).encode('utf-8')).decode('ascii')
    final_js = JS_CODE.replace('{{DASH_PAYLOAD}}', dash_payload_str)
    final_js = final_js.replace('{{PETA_CHUNKS}}', peta_b64)

    return {
        "css": CSS_CODE,
        "body_dash": dash_html_content,
        "body_peta": peta_html_content,
        "js": final_js,
        "js_dash": final_js,
        "js_peta": ""
    }


def build_combined_html(dash_payload, dash_years, peta_b64, pub_vals, jenis_vals, klu_vals, title, log):
    return build_lha_module(dash_payload, dash_years, peta_b64, pub_vals, jenis_vals, klu_vals, title, log)
