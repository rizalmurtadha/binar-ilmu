from flask import Flask, request, session, url_for, redirect, render_template, send_from_directory, make_response
from datetime import datetime
import pandas as pd
import numpy as np
import os
import pdfkit
from PyPDF2 import PdfFileMerger

import dropbox
from contextlib import closing
from io import BytesIO

app = Flask(__name__)
app.secret_key = "penilaian-sekolah-IKN"

token = "UUqaOMObW8sAAAAAAAAAAcvkmOaxYzJs7ZRhjCRDqVqvKuP-8Gd1W0n7i6CjhKNK"
dbx = dropbox.Dropbox(token)

APP_ROOT = os.path.dirname(os.path.abspath(__file__))
PATH_NILAI = os.path.join(APP_ROOT, 'nilai/')
PATH_DATA = os.path.join(APP_ROOT, 'data/')
PATH_TMP = os.path.join(APP_ROOT, 'tmp/')


def stream_dropbox_file(path):
    _,res = dbx.files_download(path)
    with closing(res) as result:
        byte_data=result.content
        return BytesIO(byte_data)

def loadDataInPd(path, source=None):
    if session["tingkat"] == "SMA" :
        path = "/SMA{}".format(path)
    
    if source == None:
        file_stream=stream_dropbox_file(path)
        return pd.read_excel(file_stream)
    else :
        if session["tingkat"] == "SMA" :
            source = "/SMA{}".format(source)

        try:
            dbx.files_get_metadata(path)
        except:
            dbx.files_copy(source, path)
        file_stream=stream_dropbox_file(path)
        return pd.read_excel(file_stream)

def saveDataToDB(pathFromHeroku, targetInDB):
    if session["tingkat"] == "SMA" :
        targetInDB = "/SMA{}".format(targetInDB) 
    with open(pathFromHeroku, 'rb') as f:
        dbx.files_upload(f.read(), targetInDB, mode=dropbox.files.WriteMode.overwrite)

def check_jenjang():
    if session['tingkat'] == "SMA":
        return "/SMA"
    else:
        return ""

def check_period():
    # define the period
    today = datetime.today()
    yr = today.year
    mo = today.month

    if (mo > 2) & (mo < 7):
        tahun_ajaran = "{}/{}".format(yr-1, yr)
        semester = 2
        folder_name_1 = "{}_{}_{}".format(yr-1, yr, semester)
    else:
        tahun_ajaran = "{}/{}".format(yr, yr+1)
        semester = 1
        folder_name_1 = "{}_{}_{}".format(yr, yr+1, semester)

    return [tahun_ajaran, semester, folder_name_1]

tahun_ajaran, semester, folder_name_1 = check_period()

def check_folder(eval_type):
    tahun_ajaran, semester, folder_name_1 = check_period()
    # create folder of year_semester if not exist
    jenjang = check_jenjang()
    try:
        dbx.files_get_metadata("{}/nilai/{}".format(jenjang, folder_name_1))
    except:
        dbx.files_create_folder("{}/nilai/{}".format(jenjang, folder_name_1))
    # create folder PTS/PAS if not exist
    try:
        dbx.files_get_metadata("{}/nilai/{}/{}".format(jenjang, folder_name_1, eval_type))
    except:
        dbx.files_create_folder("{}/nilai/{}/{}".format(jenjang, folder_name_1, eval_type))

def check_predikat(avg, sikap=False):
    if sikap==False:
        if avg >= 80:
            predikat = "Amat Baik"
        elif avg >= 70:
            predikat = "Baik"
        elif avg >= 60:
            predikat = "Cukup"
        elif avg >= 50:
            predikat = "Kurang"
        else:
            predikat = "Sangat Kurang"
    else:
        if avg == 1:
            predikat = "Kurang"
        elif avg == 2:
            predikat = "Cukup"
        elif avg == 3:
            predikat = "Baik"
        else:
            predikat = "Sangat Baik"
    return predikat

def update_kelas(nm, kls, list_nm):
    if session['tingkat'] == "SMP":
        if nm in list_nm:
            if kls == "VII":
                new_kls = "VIII"
            elif kls == "VIII":
                new_kls = "IX"
            else:
                new_kls = "Alumni"
        else:
            new_kls = kls
    elif session['tingkat'] == "SMA":
        if nm in list_nm:
            if kls == "X":
                new_kls = "XI"
            elif kls == "XI":
                new_kls = "XII"
            else:
                new_kls = "Alumni"
        else:
            new_kls = kls
    return new_kls

@app.route("/tingkat",methods=["GET", "POST"])
def tingkat():
    try:
        tingkat = request.form['tingkat']
        session["tingkat"] = tingkat
    finally:
        return redirect(url_for("login"))

@app.route("/",methods=["GET", "POST"])
def login():
    if "user" in session:
        if session["user"] == 100:
            return redirect(url_for("admin"))
        else:
            return redirect(url_for("role"))
    else:
        if request.method=="POST":
            try:
                Login = request.form['Login']
            except:
                Login = "0"
            if(Login=="1"):
                # misal input id_guru di-assign sebagai variable "id_guru"
                # misal input password di-assign sebagai variable "input_passwd"
                id_guru = int(request.form['id_guru'])
                input_passwd = request.form['password']

                # Verifikasi Admin
                if id_guru == 100 and input_passwd=="adminsmpbinarilmu":
                    session["nama_user"] = "Admin"
                    session["user"] = id_guru
                    return redirect(url_for("admin"))

                # load data guru
                data_guru = loadDataInPd("/data_guru.xlsx")

                # verifikasi akun
                list_ID_guru = data_guru.loc[:, "ID Guru"].tolist()
                if id_guru not in list_ID_guru:
                    # ID guru tidak terdaftar
                    return render_template("login.html",message="error")
                else:
                    passwd = data_guru[data_guru["ID Guru"]==id_guru]["Password"].values[0]
                    if input_passwd == passwd:
                        nama_guru = data_guru[data_guru["ID Guru"]==id_guru]["Nama"].values[0]
                        session["nama_user"] = nama_guru
                        session["user"] = id_guru
                        return redirect(url_for("role"))
                    else:
                        # Password salah
                        return render_template("login.html",message="pwdSalah")
        if "tingkat" not in session:
            return render_template("login_tingkat.html")
        return render_template("login.html")

@app.route("/role",methods=["GET", "POST"])
def role():
    if "user" in session:
        # 2. Pilih Role
        # load data guru
        data_guru = loadDataInPd("/data_guru.xlsx")
        # create dict ID guru: nama guru
        dict_guru = {}
        for i in range(data_guru.shape[0]):
            dict_guru[data_guru.loc[i,"ID Guru"]] = data_guru.loc[i,"Nama"]

        # load data guru mapel
        guru_mapel = loadDataInPd("/nilai/{}/guru_mapel.xlsx".format(folder_name_1), "/guru_mapel.xlsx")
        guru_mapel.set_index("Mata Pelajaran", inplace=True)
        # create list mapel and kelas
        mapel = guru_mapel.index.tolist()
        kelas = guru_mapel.columns.tolist()
        # create empty list
        mapel_list = []
        kelas_list = []
        id_guru = session['user']
        nama_guru = dict_guru[id_guru]
        # iterating over mapel and kelas
        for mpl in mapel:
            for kls in kelas:
                sel = guru_mapel.loc[mpl, kls]
                if sel == nama_guru:
                    mapel_list.append(mpl)
                    kelas_list.append(kls)
        # print mapel and kelas
        if len(mapel_list) > 0:
            for i in range(len(mapel_list)):
                print("Guru {} {}".format(mapel_list[i], kelas_list[i]))

        # load data wali kelas
        guru_wali_kelas = loadDataInPd("/nilai/{}/guru_wali_kelas.xlsx".format(folder_name_1), "/guru_wali_kelas.xlsx")
        # create a list of guru wali kelas
        wali_kelas_list = guru_wali_kelas["Wali Kelas"].tolist()
        # create empty list
        wali_list = []
        # iterating over the dataframe
        if nama_guru in wali_kelas_list:
            for i in range(guru_wali_kelas.shape[0]):
                sel = guru_wali_kelas.loc[i, "Wali Kelas"]
                if nama_guru == sel:
                    wali_list.append(guru_wali_kelas.loc[i, "Kelas"])
        # print wali kelas
        if len(wali_list) > 0:
            for wl in wali_list:
                print("Wali Kelas {}".format(wl))
        return render_template("role.html", mapel_list=mapel_list, kelas_list=kelas_list, wali_list=wali_list,
                                            jlh_mapel=len(mapel_list), jlh_wk=len(wali_list))
    else:
        return redirect(url_for("login"))

# ================== Admin ==============================
@app.route("/admin",methods=["GET", "POST"])
def admin():
    if session["user"] == 100:
        if request.method=="POST":
            pilihan = request.form['pilihan']
            return redirect(url_for(pilihan))
        return render_template("admin.html")
    else:
        return redirect(url_for("login"))

# --------------- Admin - Data Siswa ---------------------------
@app.route("/admin/data-siswa",methods=["GET", "POST"])
def data_siswa():
    if session["user"] == 100:
        if request.method=="POST":
            kelas = request.form['kelas']
            list_kelas = request.form.getlist('list_kelas[]')
            if kelas == "tambah":
                return kelas
            else:
                # 10 Halaman Kelas
                # kelas = "VII" #Input
                # load data siswa
                data_siswa = loadDataInPd("/data_siswa.xlsx")
                data_kelas = data_siswa[data_siswa["Kelas"] == kelas]
                list_nisn = data_kelas["NISN"].tolist()
                list_siswa = data_kelas["Nama"].tolist()
                list_kelas_siswa = data_kelas["Kelas"].tolist()

                # return str(list_kelas)
                return render_template("hal_kelas.html",kelas=kelas,list_kelas=list_kelas, list_nisn=list_nisn, jlh_list=len(list_nisn),
                                        list_siswa=list_siswa, list_kelas_siswa=list_kelas_siswa)
        # 9 Halaman Data Siswa
        # load data siswa
        data_siswa = loadDataInPd("/data_siswa.xlsx")
        list_kelas = data_siswa["Kelas"].unique()
        list_kelas = list(list_kelas)
        # for kelas in list_kelas:
        #     print("Kelas {}".format(kelas))
        return render_template("data_siswa.html",list_kelas=list_kelas,jlh_kelas=len(list_kelas), target="data_siswa")
    else:
        return redirect(url_for("login"))

@app.route("/admin/data-siswa/edit",methods=["GET", "POST"])
def edit_siswa():
    if session["user"] == 100:
        if request.method=="POST":
            try:
                save_edit = request.form["save_edit"]
            except:
                save_edit = "0"
            try:
                dropout = request.form["dropout"]
            except:
                dropout = "0"

            kelas_edit = request.form["kelas_siswa"]
            nisn_edit = int(request.form["nisn_siswa"])
            nama_edit = request.form["nama_siswa"]
            list_kelas = request.form.getlist('list_kelas[]')
            kelas = request.form['kelas']

            data_siswa = loadDataInPd("/data_siswa.xlsx")
            # nisn = 72100585
            # nama = "Agus Ahmad"  # <from input>
            # kelas = "VII"    # <from input>

            # Get Detail from database for edit page
            data_siswa_new = data_siswa.fillna('')
            data_siswa_new = data_siswa_new.loc[data_siswa_new['NISN'] == nisn_edit].values.flatten().tolist()
            # return str(int(data_siswa_new[4]))
            for i in range(len(data_siswa_new)):
                if (isinstance(data_siswa_new[i], float)):
                    data_siswa_new[i] = str(int(data_siswa_new[i]))

            pas_foto_name = ""
            ijazah_name = ""
            kk_name = ""
            akta_name = ""
            if(save_edit =="0" and dropout =="0"):
                pas_foto_list = dbx.files_list_folder(path="{}/dokumen/pas_foto".format(check_jenjang()))
                for i in range(len(pas_foto_list.entries)):
                    file_name = pas_foto_list.entries[i].name
                    if ( str(nisn_edit) == file_name.split(".")[0]):
                        pas_foto_name = "pas_foto_{}".format(file_name)
                        dbx.files_download_to_file("./tmp/{}".format(pas_foto_name), "{}/dokumen/pas_foto/{}".format(check_jenjang(), file_name))

                ijazah_list = dbx.files_list_folder(path="{}/dokumen/ijazah_sd".format(check_jenjang()))
                for i in range(len(ijazah_list.entries)):
                    file_name = ijazah_list.entries[i].name
                    if ( str(nisn_edit) == file_name.split(".")[0]):
                        ijazah_name = "ijazah_{}".format(file_name)
                        dbx.files_download_to_file("./tmp/{}".format(ijazah_name), "{}/dokumen/ijazah_sd/{}".format(check_jenjang(), file_name))


                kk_list = dbx.files_list_folder(path="{}/dokumen/kartu_keluarga".format(check_jenjang()))
                for i in range(len(kk_list.entries)):
                    file_name = kk_list.entries[i].name
                    if ( str(nisn_edit) == file_name.split(".")[0]):
                        kk_name = "kartu_keluarga{}".format(file_name)
                        dbx.files_download_to_file("./tmp/{}".format(kk_name), "{}/dokumen/kartu_keluarga/{}".format(check_jenjang(), file_name))

                akta_list = dbx.files_list_folder(path="{}/dokumen/akta_kelahiran".format(check_jenjang()))
                for i in range(len(akta_list.entries)):
                    file_name = akta_list.entries[i].name
                    if ( str(nisn_edit) == file_name.split(".")[0]):
                        akta_name = "akta_kelahiran_{}".format(file_name)
                        dbx.files_download_to_file("./tmp/{}".format(akta_name), "{}/dokumen/akta_kelahiran/{}".format(check_jenjang(), file_name))

            data_siswa.set_index("NISN", inplace=True)

            # return str(data_siswa_new[0])

            if save_edit=="1":
                # return nisn_edit
                # return str(kelas_edit+" "+nama_edit)
                # Tombol Edit
                data_siswa.loc[nisn_edit, "Nama"]           = nama_edit
                data_siswa.loc[nisn_edit, "Nama Panggilan"] = request.form["nama_panggilan"]
                # data_siswa.loc[nisn_edit, "NISN"]           = nisn_edit
                data_siswa.loc[nisn_edit, "Kelas"]          = kelas_edit
                data_siswa.loc[nisn_edit, "NIK"]            = str(request.form["NIK"])

                tmpt_lhr_siswa = request.form['tmpt_lhr_siswa']
                tgl_lhr_siswa  = request.form.get('tgl_lhr_siswa')
                data_siswa.loc[nisn_edit, "Tempat Lahir"]           = tmpt_lhr_siswa 
                data_siswa.loc[nisn_edit, "Tanggal Lahir"]          = tgl_lhr_siswa
                data_siswa.loc[nisn_edit, "Tempat, Tanggal Lahir"]  = "{}, {}".format(tmpt_lhr_siswa, tgl_lhr_siswa) if (tmpt_lhr_siswa != '') else ""    # request.form.get('TTL_siswa')
                data_siswa.loc[nisn_edit, "Jenis Kelamin"]          = request.form.get("jenis_kelamin")
                data_siswa.loc[nisn_edit, "Agama"]                  = request.form["agama_siswa"]
                data_siswa.loc[nisn_edit, "Status dalam Keluarga"]  = request.form["status"]
                data_siswa.loc[nisn_edit, "Anak ke"]                = request.form["anak_ke"]
                data_siswa.loc[nisn_edit, "Alamat Siswa"]           = request.form["alamat_siswa"]
                data_siswa.loc[nisn_edit, "Koordinat Bujur"]        = request.form["koor_bujur"]
                data_siswa.loc[nisn_edit, "Koordinat Lintang"]      = request.form["koor_lintang"]
                data_siswa.loc[nisn_edit, "Nomor Telepon/ hp siswa"] = str(request.form["telp_siswa"])
                data_siswa.loc[nisn_edit, "Sekolah Asal"]           = request.form["sekolah_asal"]
                data_siswa.loc[nisn_edit, "Tinggi Badan "]          = request.form["tinggi_badan"]
                data_siswa.loc[nisn_edit, "Berat Badan"]            = request.form["berat_badan"]

                data_siswa.loc[nisn_edit, "Nama Ayah"]                  = request.form["nama_ayah"]
                data_siswa.loc[nisn_edit, "NIK Ayah"]                   = str(request.form["NIK_ayah"])
                tmpt_lhr_ayah = request.form['tmpt_lhr_ayah']
                tgl_lhr_ayah  = request.form.get('tgl_lhr_ayah')
                data_siswa.loc[nisn_edit, "Tempat Lahir Ayah"]          = tmpt_lhr_ayah
                data_siswa.loc[nisn_edit, "Tanggal Lahir Ayah"]         = tgl_lhr_ayah
                data_siswa.loc[nisn_edit, "Tempat, Tanggal Lahir Ayah"] = "{}, {}".format(tmpt_lhr_ayah, tgl_lhr_ayah) if (tmpt_lhr_ayah != '') else ""   # request.form.get("TTL_ayah")
                data_siswa.loc[nisn_edit, "Agama Ayah"]                 = request.form["agama_ayah"]
                data_siswa.loc[nisn_edit, "Alamat Ayah"]                = request.form["alamat_ayah"]
                data_siswa.loc[nisn_edit, "Nomor Telepon/ HP Ayah"]     = str(request.form["telp_ayah"])
                data_siswa.loc[nisn_edit, "Pekerjaan Ayah"]             = request.form["pekerjaan_ayah"]
                data_siswa.loc[nisn_edit, "Instansi Tempat Bekerja"]    = request.form["instansi_ayah"]
                data_siswa.loc[nisn_edit, "Akumulasi Gaji Ayah dan Ibu"] = request.form.get("penghasilan")
                data_siswa.loc[nisn_edit, "Pendidikan Terakhir ayah"]   = request.form["pendidikan_ayah"]

                data_siswa.loc[nisn_edit, "Nama Ibu"]                   = request.form["nama_ibu"]
                data_siswa.loc[nisn_edit, "NIK Ibu"]                    = str(request.form["NIK_ibu"])
                tmpt_lhr_ibu = request.form['tmpt_lhr_ibu']
                tgl_lhr_ibu  = request.form.get('tgl_lhr_ibu')
                data_siswa.loc[nisn_edit, "Tempat Lahir Ibu"]           = tmpt_lhr_ibu
                data_siswa.loc[nisn_edit, "Tanggal Lahir Ibu"]          = tgl_lhr_ibu
                data_siswa.loc[nisn_edit, "Tempat, Tanggal Lahir Ibu"]  = "{}, {}".format(tmpt_lhr_ibu, tgl_lhr_ibu) if (tmpt_lhr_ibu != '') else "" # request.form.get("TTL_ibu")
                data_siswa.loc[nisn_edit, "Agama Ibu"]                  = request.form["agama_ibu"]
                data_siswa.loc[nisn_edit, "Alamat Ibu"]                 = request.form["alamat_ibu"]
                data_siswa.loc[nisn_edit, "No Tlp/ HP Ibu"]             = str(request.form["telp_ibu"])
                data_siswa.loc[nisn_edit, "Pekerjaan Ibu"]              = request.form["pekerjaan_ibu"]
                data_siswa.loc[nisn_edit, "Instansi Tempat Bekerja Ibu"] = request.form["instansi_ibu"]
                data_siswa.loc[nisn_edit, "Pendidikan Terakhir Ibu"]    = request.form["pendidikan_ibu"]

                data_siswa.loc[nisn_edit, "Jarak Rumah - Sekolah BI"]   = request.form["jarak_BI"]
                data_siswa.loc[nisn_edit, "Jarak Rumah - Sekolah TU"]   = request.form["jarak_TU"]

                data_siswa.reset_index(inplace=True)

                first_column = data_siswa.pop('Nama')
                second_column = data_siswa.pop('Nama Panggilan')

                data_siswa.insert(0, 'Nama Panggilan', second_column)
                data_siswa.insert(0, 'Nama', first_column)

                data_siswa.to_excel("./tmp/data_siswa_edit.xlsx", index=None)
                saveDataToDB("./tmp/data_siswa_edit.xlsx", "/data_siswa.xlsx")
                ## with open("./tmp/data_siswa_edit.xlsx", 'rb') as f:
                ##     dbx.files_upload(f.read(), "/data_siswa.xlsx", mode=dropbox.files.WriteMode.overwrite)

                # image file
                pas_foto = request.files["pas_foto"]
                ijazah = request.files["ijazah"]
                kk = request.files["kk"]
                akta_kelahiran = request.files["akta_kelahiran"]

                # Save foto
                if pas_foto.filename != '':
                    pas_foto_ext = pas_foto.filename.split(".")[1]
                    pas_foto_name = "pas_foto.{}".format(pas_foto_ext)

                    file_destination = "/".join([PATH_TMP, pas_foto_name])
                    pas_foto.save(file_destination)
                    saveDataToDB("./tmp/{}".format(pas_foto_name), "/dokumen/pas_foto/{}.{}".format(nisn_edit,pas_foto_ext))
                    ## with open("./tmp/{}".format(pas_foto_name), 'rb') as f:
                    ##     dbx.files_upload(f.read(), "/dokumen/pas_foto/{}.{}".format(nisn_edit,pas_foto_ext), mode=dropbox.files.WriteMode.overwrite)

                if ijazah.filename != '':
                    ijazah_ext = ijazah.filename.split(".")[1]
                    ijazah_name = "ijazah.{}".format(ijazah_ext)

                    file_destination = "/".join([PATH_TMP, ijazah_name])
                    ijazah.save(file_destination)
                    saveDataToDB("./tmp/{}".format(ijazah_name), "/dokumen/ijazah_sd/{}.{}".format(nisn_edit,ijazah_ext))
                    ## with open("./tmp/{}".format(ijazah_name), 'rb') as f:
                    ##     dbx.files_upload(f.read(), "/dokumen/ijazah_sd/{}.{}".format(nisn_edit,ijazah_ext), mode=dropbox.files.WriteMode.overwrite)

                if kk.filename != '':
                    kk_ext = kk.filename.split(".")[1]
                    kk_name = "kk.{}".format(kk_ext)

                    file_destination = "/".join([PATH_TMP, kk_name])
                    kk.save(file_destination)
                    saveDataToDB("./tmp/{}".format(kk_name), "/dokumen/kartu_keluarga/{}.{}".format(nisn_edit,kk_ext))
                    ## with open("./tmp/{}".format(kk_name), 'rb') as f:
                    ##     dbx.files_upload(f.read(), "/dokumen/kartu_keluarga/{}.{}".format(nisn_edit,kk_ext), mode=dropbox.files.WriteMode.overwrite)

                if akta_kelahiran.filename != '':
                    akta_kelahiran_ext = akta_kelahiran.filename.split(".")[1]
                    akta_kelahiran_name = "akta_kelahiran.{}".format(akta_kelahiran_ext)

                    file_destination = "/".join([PATH_TMP, akta_kelahiran_name])
                    akta_kelahiran.save(file_destination)
                    saveDataToDB("./tmp/{}".format(akta_kelahiran_name), "/dokumen/akta_kelahiran/{}.{}".format(nisn_edit,akta_kelahiran_ext))
                    ## with open("./tmp/{}".format(akta_kelahiran_name), 'rb') as f:
                    ##    dbx.files_upload(f.read(), "/dokumen/akta_kelahiran/{}.{}".format(nisn_edit,akta_kelahiran_ext), mode=dropbox.files.WriteMode.overwrite)

                return redirect(url_for("data_siswa"))
            elif dropout =="1":
                # return "DO"
                # Tombol Dropout
                data_siswa.loc[nisn_edit, "Kelas"] = "Dropout"
                data_siswa.reset_index(inplace=True)

                data_siswa.to_excel("./data/data_siswa.xlsx", index=None)
                saveDataToDB("./data/data_siswa.xlsx", "/data_siswa.xlsx")
                ## with open("./data/data_siswa.xlsx", 'rb') as f:
                ##     dbx.files_upload(f.read(), "/data_siswa.xlsx", mode=dropbox.files.WriteMode.overwrite)
                return redirect(url_for("data_siswa"))
        return render_template("edit_siswa_new.html",kelas=kelas,list_kelas=list_kelas, nisn_edit=nisn_edit,
                                        nama_edit=nama_edit, kelas_edit=kelas_edit, data_siswa_new=data_siswa_new,
                                        pas_foto_name=pas_foto_name, ijazah_name=ijazah_name,
                                        kk_name=kk_name, akta_name=akta_name)
    else:
        return redirect(url_for("login"))

@app.route("/admin/data-siswa/tambah",methods=["GET", "POST"])
def tambah_siswa():
    if session["user"] == 100:
        if request.method == "POST":
            # 12 Halaman Tambah Siswa
            # Unduh
            # sent ./data/data_siswa_template.xlsx to user

            # kelas = request.form["kelas"]
            try:
                tambah = request.form["unggah"]
                print(tambah)
                if tambah =="1":
                    bulk_upload = request.form["bulk_upload"]
                    # return str(bulk_upload)
                    if (bulk_upload == "1"):
                        #upload bulk
                        file = request.files["file_bulk"]
                        excel_name = file.filename
                        file_destination = "/".join([PATH_TMP,excel_name])
                        file.save(file_destination)

                        # # Unggah
                        # # uploaded file is located at ./tmp/data_siswa_template.xlsx to
                        new_data = pd.read_excel(file_destination)
                        data_siswa = loadDataInPd("/data_siswa.xlsx")
                        data_siswa = pd.concat([data_siswa, new_data], axis=0)

                        data_siswa.to_excel("./tmp/data_siswa_bulk.xlsx", index=None)
                        saveDataToDB("./tmp/data_siswa_bulk.xlsx", "/data_siswa.xlsx")
                        ## with open("./tmp/data_siswa_bulk.xlsx", 'rb') as f:
                        ##     dbx.files_upload(f.read(), "/data_siswa.xlsx", mode=dropbox.files.WriteMode.overwrite)
                        # return "berhasil upload bulk"
                        return redirect(url_for("data_siswa"))

                    # data excel
                    nama = request.form["nama"]
                    nama_panggilan = request.form["nama_panggilan"]
                    NISN = request.form["NISN"]

                    

                    kelas = request.form["kelas"]
                    NIK = str(request.form["NIK"])

                    tmpt_lhr_siswa = request.form["tmpt_lhr_siswa"]
                    tgl_lhr_siswa = request.form.get("tgl_lhr_siswa")
                    TTL_siswa = "{}, {}".format(tmpt_lhr_siswa, tgl_lhr_siswa)  if (tmpt_lhr_siswa != '') else ""
                    # TTL_siswa = request.form.get("TTL_siswa")
                    jenis_kelamin = request.form.get("jenis_kelamin")
                    agama_siswa = request.form["agama_siswa"]
                    status = request.form["status"]
                    anak_ke = request.form["anak_ke"]
                    alamat_siswa = request.form["alamat_siswa"]
                    koor_bujur = request.form["koor_bujur"]
                    koor_lintang = request.form["koor_lintang"]
                    telp_siswa = str(request.form["telp_siswa"])
                    sekolah_asal = request.form["sekolah_asal"]
                    tinggi_badan = request.form["tinggi_badan"]
                    berat_badan = request.form["berat_badan"]

                    nama_ayah = request.form["nama_ayah"]
                    NIK_ayah = str(request.form["NIK_ayah"])
                    tmpt_lhr_ayah = request.form["tmpt_lhr_ayah"]
                    tgl_lhr_ayah = request.form.get("tgl_lhr_ayah")
                    TTL_ayah = "{}, {}".format(tmpt_lhr_ayah, tgl_lhr_ayah) if (tmpt_lhr_ayah != '') else ""
                    # TTL_ayah = request.form.get("TTL_ayah")
                    agama_ayah = request.form["agama_ayah"]
                    alamat_ayah = request.form["alamat_ayah"]
                    telp_ayah = str(request.form["telp_ayah"])
                    pekerjaan_ayah = request.form["pekerjaan_ayah"]
                    instansi_ayah = request.form["instansi_ayah"]
                    penghasilan = request.form.get("penghasilan")
                    pendidikan_ayah = request.form["pendidikan_ayah"]

                    nama_ibu = request.form["nama_ibu"]
                    NIK_ibu = str(request.form["NIK_ibu"])
                    tmpt_lhr_ibu = request.form["tmpt_lhr_ibu"]
                    tgl_lhr_ibu = request.form.get("tgl_lhr_ibu")
                    TTL_ibu = "{}, {}".format(tmpt_lhr_ibu, tgl_lhr_ibu) if (tmpt_lhr_ibu != '') else ""
                    # TTL_ibu = request.form.get("TTL_ibu")
                    agama_ibu = request.form["agama_ibu"]
                    alamat_ibu = request.form["alamat_ibu"]
                    telp_ibu = str(request.form["telp_ibu"])
                    pekerjaan_ibu = request.form["pekerjaan_ibu"]
                    instansi_ibu = request.form["instansi_ibu"]
                    pendidikan_ibu = request.form["pendidikan_ibu"]

                    jarak_BI = request.form["jarak_BI"]
                    jarak_TU = request.form["jarak_TU"]

                    data_siswa = loadDataInPd("/data_siswa.xlsx")

                    data_siswa = data_siswa.append(
                        {   'Nama': nama,                       'Nama Panggilan': nama_panggilan,       'NISN': NISN,
                            'Kelas': kelas,                     'NIK': NIK,                             
                            'Tempat Lahir': tmpt_lhr_siswa,     'Tanggal Lahir': tgl_lhr_siswa,         'Tempat, Tanggal Lahir': TTL_siswa,
                            'Jenis Kelamin': jenis_kelamin,     'Agama': agama_siswa,                   'Status dalam Keluarga': status,
                            'Anak ke': anak_ke,                 'Alamat Siswa': alamat_siswa,           'Koordinat Bujur': koor_bujur,
                            'Koordinat Lintang': koor_lintang,  'Nomor Telepon/ hp siswa': telp_siswa,  'Sekolah Asal': sekolah_asal,
                            'Tinggi Badan ': tinggi_badan,      'Berat Badan': berat_badan,

                            'Nama Ayah': nama_ayah,                     'NIK Ayah': NIK_ayah,                       
                            'Tempat Lahir Ayah': tmpt_lhr_ayah,         'Tanggal Lahir Ayah': tgl_lhr_ayah,         'Tempat, Tanggal Lahir Ayah': TTL_ayah,     'Agama Ayah': agama_ayah,
                            'Alamat Ayah': alamat_ayah,                 'Nomor Telepon/ HP Ayah': telp_ayah,        'Pekerjaan Ayah': pekerjaan_ayah,
                            'Instansi Tempat Bekerja': instansi_ayah,   'Akumulasi Gaji Ayah dan Ibu': penghasilan, 'Pendidikan Terakhir ayah': pendidikan_ayah,

                            'Nama Ibu': nama_ibu,               'NIK Ibu': NIK_ibu,                            
                            'Tempat Lahir Ibu': tmpt_lhr_ibu,  'Tanggal Lahir Ibu': tgl_lhr_ibu,                'Tempat, Tanggal Lahir Ibu': TTL_ibu,
                            'Agama Ibu': agama_ibu,             'Alamat Ibu': alamat_ibu,                      'No Tlp/ HP Ibu': telp_ibu,
                            'Pekerjaan Ibu': pekerjaan_ibu,     'Instansi Tempat Bekerja Ibu': instansi_ibu,   'Pendidikan Terakhir Ibu': pendidikan_ibu,

                            'Jarak Rumah - Sekolah BI': jarak_BI, 'Jarak Rumah - Sekolah TU': jarak_TU

                        }, ignore_index=True)
                    # data_siswa.to_excel("./tmp/data_siswa.xlsx", index=None)
                    data_siswa.to_excel("./tmp/data_siswa_append.xlsx", index=None)

                    # -- Untuk Upload Ke Dropbox --
                    saveDataToDB("./tmp/data_siswa_append.xlsx", "/data_siswa.xlsx")
                    ## with open("./tmp/data_siswa_append.xlsx", 'rb') as f:
                    #### with open("./dev/data_siswa.xlsx", 'rb') as f:    # untuk reset data
                    ##     dbx.files_upload(f.read(), "/data_siswa.xlsx", mode=dropbox.files.WriteMode.overwrite)

                    # image file
                    pas_foto = request.files["pas_foto"]
                    ijazah = request.files["ijazah"]
                    kk = request.files["kk"]
                    akta_kelahiran = request.files["akta_kelahiran"]

                    # Save foto
                    if pas_foto.filename != '':
                        pas_foto_ext = pas_foto.filename.split(".")[1]
                        pas_foto_name = "pas_foto.{}".format(pas_foto_ext)

                        file_destination = "/".join([PATH_TMP, pas_foto_name])
                        pas_foto.save(file_destination)
                        saveDataToDB("./tmp/{}".format(pas_foto_name), "/dokumen/pas_foto/{}.{}".format(NISN,pas_foto_ext))
                        ## with open("./tmp/{}".format(pas_foto_name), 'rb') as f:
                        ##     dbx.files_upload(f.read(), "/dokumen/pas_foto/{}.{}".format(NISN,pas_foto_ext), mode=dropbox.files.WriteMode.overwrite)

                    if ijazah.filename != '':
                        ijazah_ext = ijazah.filename.split(".")[1]
                        ijazah_name = "ijazah.{}".format(ijazah_ext)

                        file_destination = "/".join([PATH_TMP, ijazah_name])
                        ijazah.save(file_destination)
                        saveDataToDB("./tmp/{}".format(ijazah_name), "/dokumen/ijazah_sd/{}.{}".format(NISN,ijazah_ext))
                        ## with open("./tmp/{}".format(ijazah_name), 'rb') as f:
                        ##     dbx.files_upload(f.read(), "/dokumen/ijazah_sd/{}.{}".format(NISN,ijazah_ext), mode=dropbox.files.WriteMode.overwrite)

                    if kk.filename != '':
                        kk_ext = kk.filename.split(".")[1]
                        kk_name = "kk.{}".format(kk_ext)

                        file_destination = "/".join([PATH_TMP, kk_name])
                        kk.save(file_destination)
                        saveDataToDB("./tmp/{}".format(kk_name), "/dokumen/kartu_keluarga/{}.{}".format(NISN,kk_ext))
                        ## with open("./tmp/{}".format(kk_name), 'rb') as f:
                        ##     dbx.files_upload(f.read(), "/dokumen/kartu_keluarga/{}.{}".format(NISN,kk_ext), mode=dropbox.files.WriteMode.overwrite)

                    if akta_kelahiran.filename != '':
                        akta_kelahiran_ext = akta_kelahiran.filename.split(".")[1]
                        akta_kelahiran_name = "akta_kelahiran.{}".format(akta_kelahiran_ext)

                        file_destination = "/".join([PATH_TMP, akta_kelahiran_name])
                        akta_kelahiran.save(file_destination)
                        saveDataToDB("./tmp/{}".format(akta_kelahiran_name), "/dokumen/akta_kelahiran/{}.{}".format(NISN,akta_kelahiran_ext))
                        ## with open("./tmp/{}".format(akta_kelahiran_name), 'rb') as f:
                        ##     dbx.files_upload(f.read(), "/dokumen/akta_kelahiran/{}.{}".format(NISN,akta_kelahiran_ext), mode=dropbox.files.WriteMode.overwrite)
                    # return "berhasil upload"
                    return redirect(url_for("data_siswa"))
            except:
                # return "masuk except, error di tambah single"
                try:
                    # return "masuk sini"
                    bulk_upload = request.form["bulk_upload"]
                    return render_template("tambah_siswa_new.html", bulk_upload=bulk_upload)
                except:
                    return redirect(url_for("data_siswa"))
        return render_template("tambah_siswa_new.html", bulk_upload="0")
    else:
        return redirect(url_for("login"))

# --------------- Admin - Data Guru -------------------------
@app.route("/admin/data-guru",methods=["GET", "POST"])
def data_guru():
    if session["user"] == 100:
        if request.method=="POST":
            pilihan = request.form['pilihan']
            return pilihan
            return render_template("role_mapel.html", nama_mapel=nama_mapel, kelas=kelas)
        # 13 Halaman Data Guru
        # load data guru
        data_guru = loadDataInPd("/data_guru.xlsx")
        # data_guru.to_excel("./data/data_guru.xlsx", index=None)
        # data_guru.head()
        list_id = data_guru["ID Guru"].tolist()
        list_nama = data_guru["Nama"].tolist()
        list_password = data_guru["Password"].tolist()

        return render_template("data_guru.html", jlh_list=len(list_id), list_id=list_id,
                                    list_nama=list_nama, list_password=list_password)
    else:
        return redirect(url_for("login"))

@app.route("/admin/data-guru/edit",methods=["GET", "POST"])
def edit_guru():
    if session["user"] == 100:
        if request.method=="POST":
            try:
                save_edit = request.form["save_edit"]
            except:
                save_edit = "0"

            id_sel = int(request.form["id_sel"])    # id_guru = 101
            nama_sel = request.form["nama_sel"]     # nama = "Wahyu Mediana"  # <from input>
            pass_sel = request.form["pass_sel"]     # password = "smpbinarilmu" # <from input>

            if save_edit=="1":
                # return str(nama_sel+" "+pass_sel)
                # 14 Halaman Edit Data Guru
                # load data guru
                data_guru = loadDataInPd("/data_guru.xlsx")

                data_guru.set_index("ID Guru", inplace=True)
                data_guru.loc[id_sel, "Nama"] = nama_sel
                data_guru.loc[id_sel, "Password"] = pass_sel
                data_guru.reset_index(inplace=True)

                data_guru.to_excel("./data/data_guru.xlsx", index=None)
                saveDataToDB("./data/data_guru.xlsx", "/data_guru.xlsx")
                ## with open("./data/data_guru.xlsx", 'rb') as f:
                ##     dbx.files_upload(f.read(), "/data_guru.xlsx", mode=dropbox.files.WriteMode.overwrite)
                return redirect(url_for("data_guru"))
            return render_template("edit_guru.html", id_sel=id_sel, nama_sel=nama_sel, pass_sel=pass_sel)
    else:
        return redirect(url_for("login"))

@app.route("/admin/data-guru/tambah",methods=["GET", "POST"])
def tambah_guru():
    if session["user"] == 100:
        if request.method == "POST":
            pass
            try:
                tambah = request.form["tambah"]
                name = request.form["nama"]         # name = "Isman Kurniawan"   # <from input>
            except:
                tambah = "0"
            # 15 Halaman Tambah Data Guru
            # load data guru
            data_guru = loadDataInPd("/data_guru.xlsx")
            id_now = data_guru["ID Guru"].tolist()[-1] + 1
            if tambah == "1":
                data_guru = data_guru.append({'ID Guru': id_now, 'Nama': name, 'Password': 'smpbinarilmu'}, ignore_index=True)
                data_guru.to_excel("./data/data_guru.xlsx", index=None)
                saveDataToDB("./data/data_guru.xlsx", "/data_guru.xlsx")
                ## with open("./data/data_guru.xlsx", 'rb') as f:
                ##    dbx.files_upload(f.read(), "/data_guru.xlsx", mode=dropbox.files.WriteMode.overwrite)
                return redirect(url_for("data_guru"))
            return render_template("tambah_guru.html",id_now=id_now)
    else:
        return redirect(url_for("login"))

# --------------- Plot Pengajar ----------------------
@app.route("/admin/plotting-pengajaran",methods=["GET", "POST"])
def plot_pengajaran():
    if session["user"] == 100:
        # 16 Halaman Plotting Pengajaran
        # load guru mapel
        guru_mapel = loadDataInPd("/guru_mapel.xlsx")
        # buat cek data
        # guru_mapel.to_excel("./data/guru_mapel.xlsx", index=None)
        # return "GG"
        list_mapel = guru_mapel["Mata Pelajaran"].tolist()
        list_kelasVII = guru_mapel["VII"].tolist()
        list_kelasVIII = guru_mapel["VIII"].tolist()
        list_kelasIX = guru_mapel["IX"].tolist()

        if request.method=="POST":
            list_VII_baru = request.form.getlist('list_kelasVII[]')
            list_VIII_baru = request.form.getlist('list_kelasVIII[]')
            list_IX_baru = request.form.getlist('list_kelasIX[]')
            newData = pd.DataFrame(list(zip(list_mapel, list_VII_baru,list_VIII_baru,list_IX_baru)), columns =['Mata Pelajaran', 'VII', 'VIII', 'IX'])

            newData.to_excel("./data/guru_mapel.xlsx", index=None)
            # update guru_mapel.xlsx according to the input
            # guru_mapel.to_excel("./data/guru_mapel.xlsx", index=None)
            saveDataToDB("./data/guru_mapel.xlsx", "/guru_mapel.xlsx")
            ## with open("./data/guru_mapel.xlsx", 'rb') as f:
            ##     dbx.files_upload(f.read(), "/guru_mapel.xlsx", mode=dropbox.files.WriteMode.overwrite)
            return redirect(url_for("plot_pengajaran"))
        # get list guru
        # load data guru
        data_guru = loadDataInPd("/data_guru.xlsx")
        list_guru = data_guru["Nama"].tolist()
        return render_template("plot_pengajaran.html",list_mapel=list_mapel,jlh_list=len(list_mapel),list_guru=list_guru,
                                    list_kelasVII=list_kelasVII, list_kelasVIII=list_kelasVIII,list_kelasIX=list_kelasIX,jlh_guru=len(list_guru))
    else:
        return redirect(url_for("login"))

# ------------------ Plot Wali --------------------------
@app.route("/admin/plotting-wali",methods=["GET", "POST"])
def plot_wali():
    if session["user"] == 100:
        # 16 Halaman Plotting Pengajaran
        # load guru mapel
        guru_wali_kelas = loadDataInPd("/guru_wali_kelas.xlsx")
        # buat cek data
        guru_wali_kelas.to_excel("./data/guru_wali_kelas.xlsx", index=None)
        # return "test plot wali"
        list_wali = guru_wali_kelas["Wali Kelas"].tolist()
        list_kelas = guru_wali_kelas["Kelas"].tolist()
        # get list guru
        # load data guru
        data_guru = loadDataInPd("/data_guru.xlsx")
        list_guru = data_guru["Nama"].tolist()
        if request.method=="POST":
            list_wali_baru = request.form.getlist('list_wali[]')
            newData = pd.DataFrame(list(zip(list_kelas, list_wali_baru)), columns =['Kelas', 'Wali Kelas'])
            newData.to_excel("./data/guru_wali_kelas_baru.xlsx", index=None)
            # update guru_mapel.xlsx according to the input
            ### guru_mapel.to_excel("./data/guru_mapel.xlsx", index=None)
            saveDataToDB("./data/guru_wali_kelas_baru.xlsx", "/guru_wali_kelas.xlsx")
            ## with open("./data/guru_wali_kelas_baru.xlsx", 'rb') as f:
            ##     dbx.files_upload(f.read(), "/guru_wali_kelas.xlsx", mode=dropbox.files.WriteMode.overwrite)
            return redirect(url_for("plot_wali"))
        return render_template("plot_wali.html",jlh_list=len(list_kelas),list_guru=list_guru, jlh_guru=len(list_guru),list_kelas=list_kelas,list_wali=list_wali)
    else:
        return redirect(url_for("login"))

# ---------- Update Kelas / Kelulusan -------------------
@app.route("/admin/update-kelas",methods=["GET", "POST"])
def update_kelas_kelulusan():
    if session["user"] == 100:
        if request.method=="POST":
            kelas = request.form['kelas']           # kelas = "VII" #Input
            list_kelas = request.form.getlist('list_kelas[]')
            try:
                update = request.form['update_kelas']
                checked_siswa = request.form.getlist('chkbox[]')        # checked_siswa = ["Agus Ahmad", "Ajat Wahyudin"]
            except:
                update = "0"
                checked_siswa = []

            if update == "1":
                data_siswa = loadDataInPd("/nilai/{}/data_siswa.xlsx".format(folder_name_1), "/data_siswa.xlsx")                        
                for i in range(data_siswa.shape[0]):
                    nm = data_siswa.loc[i,"Nama"]
                    kls = data_siswa.loc[i,"Kelas"]
                    new_kls = update_kelas(nm, kls, checked_siswa)
                    data_siswa.loc[i,"Kelas"] = new_kls
                data_siswa.to_excel("./data/data_siswa_updated.xlsx", index=None)
                # save updated data_siswa.xlsx to DB
                saveDataToDB("./data/data_siswa_updated.xlsx", "/data_siswa.xlsx")
                ## with open("./data/data_siswa_updated.xlsx", 'rb') as f:    # untuk reset data
                ##     dbx.files_upload(f.read(), "/data_siswa.xlsx", mode=dropbox.files.WriteMode.overwrite)
                
                # cek if there is "Alumni" in data_siswa
                # data_siswa = pd.read_excel("tmp/data_siswa_append.xlsx")        # Testing
                data_siswa = loadDataInPd("/data_siswa.xlsx")
                if data_siswa.loc[data_siswa['Kelas'] == "Alumni"].shape[0] != 0:
                    alumni = loadDataInPd("/data_alumni.xlsx")  
                    alumni = alumni.append(data_siswa.loc[data_siswa['Kelas'] == "Alumni"], ignore_index=True)
                    data_siswa = data_siswa.loc[data_siswa['Kelas'] != "Alumni"]

                    data_siswa.to_excel("data/data_siswa_updated.xlsx", index=None)
                    alumni.to_excel("data/data_alumni_updated.xlsx", index=None)

                    # Save updated data_siswa & data_alumni to DB 
                    saveDataToDB("./data/data_siswa_updated.xlsx", "/data_siswa.xlsx")
                    ## with open("./data/data_siswa_updated.xlsx", 'rb') as f:    # untuk reset data
                    ##     dbx.files_upload(f.read(), "/data_siswa.xlsx", mode=dropbox.files.WriteMode.overwrite)

                    saveDataToDB("./data/data_alumni_updated.xlsx", "/data_alumni.xlsx")
                    ## with open("./data/data_alumni_updated.xlsx", 'rb') as f:    # untuk reset data
                    ##     dbx.files_upload(f.read(), "/data_alumni.xlsx", mode=dropbox.files.WriteMode.overwrite)
                return redirect(url_for("update_kelas_kelulusan"))
            else:
                # 10 Halaman Kelas
                data_siswa = loadDataInPd("/data_siswa.xlsx")
                data_kelas = data_siswa[data_siswa["Kelas"] == kelas]       
                list_nisn = data_kelas["NISN"].tolist()
                list_siswa = data_kelas["Nama"].tolist()
                list_kelas_siswa = data_kelas["Kelas"].tolist()
                return render_template("update_kelas.html",kelas=kelas,list_kelas=list_kelas, list_nisn=list_nisn, jlh_list=len(list_nisn), list_siswa=list_siswa, list_kelas_siswa=list_kelas_siswa)
        # 9 Halaman Data Siswa
        # load data siswa
        data_siswa = loadDataInPd("/data_siswa.xlsx")
        list_kelas = data_siswa["Kelas"].unique()
        list_kelas = list(list_kelas)
        # for kelas in list_kelas:
        #     print("Kelas {}".format(kelas))
        return render_template("data_siswa.html",list_kelas=list_kelas,jlh_kelas=len(list_kelas), target="update_kelas_kelulusan")
    else:
        return redirect(url_for("login"))

# ------------------- Update Mapel ---------------------
@app.route("/admin/update-mata-pelajaran",methods=["GET", "POST"])
def update_mapel():
    if session["user"] == 100:
        list_mapel = loadDataInPd("/list_mapel.xlsx")
        list_mapel.to_excel("./tmp/list_mapel.xlsx", index=None)
        if request.method=="POST":
            try:
                unggah = request.form['unggah']
            except:
                unggah = "0"
            if unggah=="1":
                file = request.files["file"]
                file_destination = "/".join([PATH_TMP,"list_mapel.xlsx"])
                file.save(file_destination)
                # Save list mapel new to DB
                saveDataToDB("./tmp/list_mapel.xlsx", "/list_mapel.xlsx")
                ## with open("./tmp/list_mapel.xlsx", 'rb') as f:
                ##    dbx.files_upload(f.read(), "/list_mapel.xlsx", mode=dropbox.files.WriteMode.overwrite)

                # Change file guru_mapel.xlsx
                # load file “list_mapel.xlsx”
                list_mapel = loadDataInPd("/list_mapel.xlsx")
                # list_mapel = pd.read_excel('tmp/list_mapel.xlsx')

                # create new column
                list_mapel["VII"] = ""
                list_mapel["VIII"] = ""
                list_mapel["IX"] = ""

                # save as new “guru_mapel.xlsx”
                list_mapel.to_excel("./data/guru_mapel.xlsx", index=None)
                saveDataToDB("./data/guru_mapel.xlsx", "/guru_mapel.xlsx")
                ## with open("./data/guru_mapel.xlsx", 'rb') as f:
                ##     dbx.files_upload(f.read(), "/guru_mapel.xlsx", mode=dropbox.files.WriteMode.overwrite)
                return redirect(url_for("plot_pengajaran"))
        return render_template("update_mapel.html")
    else:
        return redirect(url_for("login"))



# ================== Mapel ==============================
@app.route("/mapel",methods=["GET", "POST"])
def role_mapel():
    if "user" in session:
        if request.method=="POST":
            nama_mapel = request.form['mapel']
            kelas = request.form['kelas']
            return render_template("role_mapel.html", nama_mapel=nama_mapel, kelas=kelas)
        return redirect(url_for("role"))
    else:
        return redirect(url_for("login"))

@app.route("/mapel/aspek-materi",methods=["GET", "POST"])
def role_mapel_menu():
    if "user" in session:
        if request.method=="POST":
            # 3. Menu role guru MK
            # input from previous step
            eval_type = request.form['eval_type']   #"PTS"
            pelajaran = request.form['mapel']       #"Matematika"
            kelas = request.form['kelas']           #"VII"
            try:
                # Reset Request from rekap mapel menu
                status = int(request.form['status'])
                tahun_ajaran, semester, folder_name_1 = check_period()
                if status == 0:
                    try:
                        # delete form nilai & komentar
                        jenjang = check_jenjang()
                        dbx.files_delete("{}/nilai/{}/{}/form_nilai_{}_{}.xlsx".format(jenjang, folder_name_1, eval_type, pelajaran, kelas))
                        dbx.files_delete("{}/nilai/{}/{}/Komentar_{}_{}.xlsx".format(jenjang, folder_name_1, eval_type, pelajaran, kelas))
                        # load status file
                        status_nilai = loadDataInPd("/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type))
                        status_nilai.set_index("Mata Pelajaran", inplace=True)
                        status_nilai.loc[pelajaran, kelas] = 0
                        status_nilai.to_excel("./tmp/status_nilai.xlsx")
                        saveDataToDB("./tmp/status_nilai.xlsx", "/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type))
                        ## with open("./tmp/status_nilai.xlsx", 'rb') as f:
                        ##     dbx.files_upload(f.read(), "/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type), mode=dropbox.files.WriteMode.overwrite)
                    except:
                        print("gagal hapus data mapel")
            except:
                status = None
            # check input status
            check_folder(eval_type)
            tahun_ajaran, semester, folder_name_1 = check_period()

            # copy template status nilai if not exist
            try:
                dbx.files_get_metadata("{}/nilai/{}/{}/status_nilai.xlsx".format(check_jenjang(), folder_name_1, eval_type))
            except:
                status_nilai = loadDataInPd("/nilai/{}/guru_mapel.xlsx".format(folder_name_1), "/guru_mapel.xlsx")
                status_nilai.iloc[:,1:] = 0

                status_nilai.to_excel("./data/status_nilai.xlsx", index=None)
                saveDataToDB("./data/status_nilai.xlsx", "/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type))
                ## with open("./data/status_nilai.xlsx", 'rb') as f:    
                ##     dbx.files_upload(f.read(), "/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type), mode=dropbox.files.WriteMode.overwrite)

            # load status nilai
            status_nilai = loadDataInPd("/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type))
            status_nilai.set_index("Mata Pelajaran", inplace=True)

            if status == None:
                status = status_nilai.loc[pelajaran, kelas]
            if status == 0:
                # go to input nilai
                return render_template("aspek_materi.html", eval_type=eval_type, pelajaran=pelajaran, kelas=kelas)
            else:
                # go to rekap page
                return redirect(url_for("mapel_rekap", eval_type=eval_type, pelajaran=pelajaran, kelas=kelas))
        else:
            return redirect(url_for("role"))
    else:
        return redirect(url_for("login"))

@app.route("/input",methods=["GET", "POST"])
def menu_input():
    if "user" in session:
        if request.method=="POST":
            # 4. Input aspek materi
            aspek_materi = request.form['aspek_materi']
            eval_type = request.form['eval_type']   #"PTS"
            pelajaran = request.form['pelajaran']   #"Matematika"
            kelas = request.form['kelas']           #"VII"
            try:
                create = request.form['create']
            except:
                create = "0"
            
            if create =="1":
                save_template_form_nilai(pelajaran, kelas, aspek_materi,eval_type)
                create = "0"

            try:
                unggah = request.form['unggah']
            except:
                unggah = "0"

            if unggah=="1":
                file = request.files["file"]
                aspek_materi_list = getAspekMateriList(aspek_materi)
                flag, msg = unggah_form_nilai(file, eval_type, pelajaran, kelas, aspek_materi_list)
                if flag == 1:
                    return render_template("unggah_nilai.html", aspek_materi=aspek_materi, eval_type=eval_type, pelajaran=pelajaran, kelas=kelas, msg=msg)
                return redirect(url_for("mapel_rekap", eval_type=eval_type, pelajaran=pelajaran, kelas=kelas, aspek_materi=aspek_materi))
            return render_template("unggah_nilai.html", aspek_materi=aspek_materi, eval_type=eval_type, pelajaran=pelajaran, kelas=kelas )
        else:
            return redirect(url_for("role"))
    else:
        return redirect(url_for("login"))

@app.route("/rekap-mapel",methods=["GET", "POST"])
def mapel_rekap():
    if "user" in session:
        if request.method=="GET":
            eval_type=request.args.get('eval_type')
            pelajaran=request.args.get('pelajaran')
            kelas=request.args.get('kelas')
            nisn_siswa = None
            save_komentar ="0"
            # aspek_materi=request.args.get('aspek_materi')
        cetak_semua = "0"
        lihat = None
        if request.method=="POST":
            eval_type = request.form['eval_type']   #"PTS"
            pelajaran = request.form['pelajaran']   #"Matematika"
            kelas = request.form['kelas']           #"VII"
            # cek save komentar
            try:
                save_komentar = request.form['save_komentar']
            except:
                save_komentar="0"
            # Tombol Lihat
            try:
                lihat = request.form['lihat']
                nisn_siswa = int(request.form['nisn_siswa']) # ex:72100585  obtained from previous step
            except:
                lihat="0"
                nisn_siswa = None
            # Tombol Cetak Semua
            cetak_semua = request.form['cetak_semua']

        # 6. Tampilan rekap siswa
        # identitas rekap
        tahun_ajaran, semester, folder_name_1 = check_period()
        # load form nilai
        nilai_siswa = loadDataInPd("/nilai/{}/{}/form_nilai_{}_{}.xlsx".format(folder_name_1, eval_type, pelajaran, kelas))
        # show the list of siswa
        list_nisn = nilai_siswa["NISN"].tolist()
        list_siswa = nilai_siswa["Nama"].tolist()
        # load form komentar
        form_komentar = loadDataInPd("/nilai/{}/{}/Komentar_{}_{}.xlsx".format(folder_name_1, eval_type, pelajaran, kelas))
        form_komentar.set_index("NISN", inplace=True)

        if save_komentar=="1":
            #  update : dari nilai_mapel view -> Tombol Submit Komentar
            komentar = request.form['komentar']     # "Progres siswa sudah baik" # obtained from text box
            form_komentar.loc[(nisn_siswa), "Komentar"] = komentar
            # save komentar dataframe
            form_komentar.reset_index(inplace=True)
            form_komentar.to_excel("./nilai/Komentar_{}_{}.xlsx".format(pelajaran, kelas), index=None)
            saveDataToDB("./nilai/Komentar_{}_{}.xlsx".format(pelajaran, kelas), "/nilai/{}/{}/Komentar_{}_{}.xlsx".format(folder_name_1, eval_type, pelajaran, kelas))
            ## with open("./nilai/Komentar_{}_{}.xlsx".format(pelajaran, kelas), 'rb') as f:
            ##     dbx.files_upload(f.read(), "/nilai/{}/{}/Komentar_{}_{}.xlsx".format(folder_name_1, eval_type, pelajaran, kelas), mode=dropbox.files.WriteMode.overwrite)

        if nisn_siswa != None and save_komentar !="1":
            # present nilai siswa : rekap_mapel View -> Tombol lihat/cetak
            nilai_siswa_sel = nilai_siswa[nilai_siswa["NISN"] == nisn_siswa]
            aspek_materi = list(nilai_siswa_sel)

            komentar_siswa_sel = form_komentar.loc[nisn_siswa,"Komentar"]
            try:
                char_count = len(komentar_siswa_sel)
            except:
                char_count = 0
            if pd.isna(komentar_siswa_sel):
                komentar_siswa_sel = ""
            # return str((komentar))
            batas = len(aspek_materi) - 3
            aspek_materi = [aspek_materi[i] for i in range(6,batas) ]
            aspek_materi = aspek_materi[0::2]
            aspek_materi = [aspek_materi[i].replace("_Pengetahuan","") for i in range(len(aspek_materi)) ]
            nilai_sel = []

            for i in range(nilai_siswa_sel.shape[1]):
                if i > batas :
                    nilai_sel.append("{0:.2f}".format(nilai_siswa_sel.iloc[0,i]))
                else:
                    nilai_sel.append(nilai_siswa_sel.iloc[0,i])
            # return str(nilai_sel)
            html= render_template("nilai_mapel.html",nilai_sel=nilai_sel, tahun_ajaran=tahun_ajaran, cetak="0",lihat=lihat, nisn_siswa=nisn_siswa,
                                        pelajaran=pelajaran, kelas=kelas, semester=semester, aspek_materi=aspek_materi, komentar_siswa_sel=komentar_siswa_sel,char_count=char_count,
                                        eval_type=eval_type, nama_guru=session['nama_user'],jlh_aspek=len(aspek_materi))

        if cetak_semua == "1":
            # Cetak Semua >> iterating to print the data of all student (./nilai/Name_NISN_Mapel?.pdf)
            pdfs = []
            for siswa in range(len(list_siswa)):
                nilai_siswa_sel = nilai_siswa[nilai_siswa["NISN"] == int(list_nisn[siswa])]
                aspek_materi = list(nilai_siswa_sel)
                komentar_siswa_sel = form_komentar.loc[(list_nisn[siswa]),"Komentar"]
                try:
                    char_count = len(komentar_siswa_sel)
                except:
                    char_count = 0
                if pd.isna(komentar_siswa_sel):
                    komentar_siswa_sel = ""
                # return str((komentar))
                batas = len(aspek_materi) - 3
                aspek_materi = [aspek_materi[i] for i in range(6,batas) ]
                aspek_materi = aspek_materi[0::2]
                aspek_materi = [aspek_materi[i].replace("_Pengetahuan","") for i in range(len(aspek_materi)) ]
                nilai_sel = []

                for i in range(nilai_siswa_sel.shape[1]):
                    if i > batas :
                        nilai_sel.append("{0:.2f}".format(nilai_siswa_sel.iloc[0,i]))
                    else:
                        nilai_sel.append(nilai_siswa_sel.iloc[0,i])

                html= render_template("nilai_mapel.html",nilai_sel=nilai_sel, tahun_ajaran=tahun_ajaran, cetak="0",lihat=lihat, nisn_siswa=nisn_siswa,
                                            pelajaran=pelajaran, kelas=kelas, semester=semester, aspek_materi=aspek_materi, komentar_siswa_sel=komentar_siswa_sel,char_count=char_count,
                                            eval_type=eval_type, nama_guru=session['nama_user'],jlh_aspek=len(aspek_materi))

                filename_pdf = "./nilai/{}_{}_{}.pdf".format(list_siswa[siswa], list_nisn[siswa], pelajaran)
                pdf = createPdf(html, filename_pdf)
                pdfs.append(filename_pdf)
                # End of loop
            # merging pdf file
            merger = PdfFileMerger()
            for pdf in pdfs:
                merger.append(pdf,import_bookmarks=False)
            merger.write("nilai/Rekap_Nilai_{}_{}.pdf".format(pelajaran, kelas))
            merger.close()
            filenames = "Rekap_Nilai_{}_{}.pdf".format(pelajaran, kelas)
            return redirect(url_for('download_from_directory', source='nilai',filename=filenames))
        elif cetak_semua =="0":
            if lihat == "0":
                filename_pdf = "{}_{}_{}.pdf".format(nilai_sel[1], nilai_sel[0], pelajaran)
                pdf = createPdf(html)
                response = createResponse(pdf, filename_pdf)
                return response
            elif lihat == "1":
                return html
        return render_template("rekap_mapel.html",jlh_list=len(list_siswa), list_nisn=list_nisn, list_siswa=list_siswa,
                                        eval_type=eval_type, pelajaran=pelajaran, kelas=kelas)
    else:
        return redirect(url_for("login"))

# ================ Wali Kelas ===========================
@app.route("/wali-kelas",methods=["GET", "POST"])
def role_wali():
    if "user" in session:
        if request.method=="POST":
            nama_mapel = request.form['mapel']
            kelas = request.form['kelas']
            return render_template("role_wali.html", nama_mapel=nama_mapel, kelas=kelas)
        return redirect(url_for("role"))
    else:
        return redirect(url_for("login"))

@app.route("/wali-kelas/tipe-semester",methods=["GET", "POST"])
def role_wali_menu():
    if "user" in session:
        if request.method=="POST":
            eval_type = request.form['eval_type']   #"PTS"
            pelajaran = request.form['mapel']       #"Wali Kelas"
            kelas = request.form['kelas']           #"VII"
            try:
                if request.form['generate'] =="1":
                    generate= True
            except:
                generate = None
            # 7. Menu role wali kelas
            # info from previous step
            tahun_ajaran, semester, folder_name_1 = check_period()
            # check the existence of folder
            check_folder(eval_type)
            # check the existence of rekap file
            if generate == None:
                try:
                    dbx.files_get_metadata("{}/nilai/{}/{}/Rekap_Nilai_{}".format(check_jenjang(), folder_name_1, eval_type, kelas))
                    generate = False
                except:
                    generate = True
            # generate rekap nilai
            if generate:
                # load list_mapel
                guru_mapel = loadDataInPd("/nilai/{}/guru_mapel.xlsx".format(folder_name_1), "/guru_mapel.xlsx")
                guru_mapel.set_index("Mata Pelajaran", inplace=True)
                # create list mapel and kelas
                mapel = guru_mapel.index.tolist()
                # create list of exist file in folder nilai/tahun_sem/PTS or PAS
                # a = dbx.files_list_folder(path="{}/nilai/{}/{}".format(check_jenjang(), folder_name_1, eval_type, kelas))
                a = dbx.files_list_folder(path="{}/nilai/{}/{}".format(check_jenjang(), folder_name_1, eval_type))
                file_list = []
                for i in range(len(a.entries)):
                    file_name = a.entries[i].name
                    file_list.append(file_name.split(".")[0])
                # print(file_list)
                # create file rekap nilai
                # load data siswa
                data_siswa = loadDataInPd("/nilai/{}/data_siswa.xlsx".format(folder_name_1), "/data_siswa.xlsx")
                form_nilai = data_siswa[data_siswa["Kelas"] == kelas][["NISN", "Nama"]]
                form_nilai.reset_index(inplace=True, drop=True)
                for mpl in mapel:
                    for aspek in ["Sikap", "Keterampilan", "Pengetahuan"]:
                        form_nilai["{}_{}".format(mpl, aspek)] = ""
                    file_name = "form_nilai_{}_{}".format(mpl, kelas)
                    if file_name in file_list:
                        # load file name
                        form_nilai_mapel = loadDataInPd("/nilai/{}/{}/{}.xlsx".format(folder_name_1, eval_type, file_name))
                        nilai_sikap = form_nilai_mapel["Spiritual_Predikat"] + form_nilai_mapel["Sosial_Predikat"]
                        tmp = round(nilai_sikap/2,0)
                        form_nilai["{}_Sikap".format(mpl)] = tmp
                        form_nilai["{}_Pengetahuan".format(mpl)] = form_nilai_mapel["Nilai Akhir Pengetahuan"].values
                        form_nilai["{}_Keterampilan".format(mpl)] = form_nilai_mapel["Nilai Akhir Keterampilan"].values
                aspek_dict = {}; pred_dict = {}
                for i in ["Sikap", "Pengetahuan", "Keterampilan"]:
                    aspek_dict[i] = []
                    pred_dict[i] = []
                form_nilai.to_excel("tmp/form_nilai_raport.xlsx")
                for j in range(form_nilai.shape[0]):
                    for i in ["Sikap", "Pengetahuan", "Keterampilan"]:
                        tmp = []
                        for mpl in mapel:
                            nilai = form_nilai.loc[j, "{}_{}".format(mpl, i)]
                            if type(nilai) != str:
                                tmp.append(nilai)
                        tmp = np.average(tmp)
                        if i == "Sikap":
                            sikap = True
                            tmp = round(tmp,0)
                        else:
                            sikap = False
                        aspek_dict[i].append(tmp)
                        pred_dict[i].append(check_predikat(tmp, sikap=sikap))
                for i in ["Sikap", "Pengetahuan", "Keterampilan"]:
                    form_nilai["{}_Avg".format(i)] = aspek_dict[i]
                    form_nilai["{}_Pred".format(i)] = pred_dict[i]
            #             for aspek in ["Sikap", "Keterampilan", "Pengetahuan"]:
            #                 form_nilai["{}_{}".format(mpl, aspek)] = form_nilai_mapel[aspek].values
            #     form_nilai["Rata_rata"] = form_nilai.iloc[:,1:].mean(axis=1).values
            #     form_nilai["Predikat"] = form_nilai["Rata_rata"].apply(lambda x: check_predikat(x))
                form_nilai.to_excel("./nilai/Rekap_Nilai_{}.xlsx".format(kelas), index=None)
                saveDataToDB("./nilai/Rekap_Nilai_{}.xlsx".format(kelas), "/nilai/{}/{}/Rekap_Nilai_{}.xlsx".format(folder_name_1, eval_type, kelas))
                ## with open("./nilai/Rekap_Nilai_{}.xlsx".format(kelas), 'rb') as f:
                ##     dbx.files_upload(f.read(), "/nilai/{}/{}/Rekap_Nilai_{}.xlsx".format(folder_name_1, eval_type, kelas), mode=dropbox.files.WriteMode.overwrite)
                
                # create file komentar
                # create komentar dataframe
                komentar = 0
                try:
                    dbx.files_get_metadata("{}/nilai/{}/{}/Komentar_{}.xlsx".format(check_jenjang(), folder_name_1, eval_type, kelas))
                except:
                    komentar = 1
                if komentar:
                    form_komentar = data_siswa[data_siswa["Kelas"] == kelas][["NISN", "Nama"]]
                    form_komentar["Komentar"] = ""
                    form_komentar.to_excel("./nilai/Komentar_{}.xlsx".format(kelas), index=None)
                    saveDataToDB("./nilai/Komentar_{}.xlsx".format(kelas), "/nilai/{}/{}/Komentar_{}.xlsx".format(folder_name_1, eval_type, kelas))
                    ## with open("./nilai/Komentar_{}.xlsx".format(kelas), 'rb') as f:
                    ##     dbx.files_upload(f.read(), "/nilai/{}/{}/Komentar_{}.xlsx".format(folder_name_1, eval_type, kelas), mode=dropbox.files.WriteMode.overwrite)
                return redirect(url_for("wali_rekap",eval_type=eval_type, pelajaran=pelajaran, kelas=kelas))
        else:
            return redirect(url_for("role"))
    else:
        return redirect(url_for("login"))

@app.route("/rekap-wali",methods=["GET", "POST"])
def wali_rekap():
    if "user" in session:
        if request.method=="GET":
            eval_type=request.args.get('eval_type')
            pelajaran=request.args.get('pelajaran')
            kelas=request.args.get('kelas')
            nisn_siswa = None
            save_komentar ="0"
        cetak_semua = "0"
        update = "0"
        lihat = None
        if request.method=="POST":
            eval_type = request.form['eval_type']   #"PTS"
            pelajaran = request.form['pelajaran']   #"Matematika"
            kelas = request.form['kelas']       #"VII"
            # cek save komentar
            try:
                save_komentar = request.form['save_komentar']
            except:
                save_komentar="0"
            # Tombol Lihat
            try:
                lihat = request.form['lihat']
                nisn_siswa = int(request.form['nisn_siswa']) # 72100585 obtained from previous step
            except:
                # lihat="0"
                nisn_siswa = None
            try:
                update = request.form['update_kelas']
            except:
                update = "0"
            # checkbox
            try:
                checked_siswa = request.form.getlist('chkbox[]')
            except:
                checked_siswa = []
            # Tombol Cetak Semua
            cetak_semua = request.form['cetak_semua']


        # tampilan rekap per siswa
        tahun_ajaran, semester, folder_name_1 = check_period()
        # load form nilai
        form_nilai = loadDataInPd("/nilai/{}/{}/Rekap_Nilai_{}.xlsx".format(folder_name_1, eval_type, kelas))

        list_nisn = form_nilai["NISN"].tolist()
        list_siswa = form_nilai["Nama"].tolist()

        form_nilai.set_index("NISN", inplace=True)
        # load form komentar
        form_komentar = loadDataInPd("/nilai/{}/{}/Komentar_{}.xlsx".format(folder_name_1, eval_type, kelas))
        form_komentar.set_index("NISN", inplace=True)
        # show the list of siswa

        if update == "1":
            # return str(len(checked_siswa))
            # update kelas
            data_siswa = loadDataInPd("/nilai/{}/data_siswa.xlsx".format(folder_name_1), "/data_siswa.xlsx")
            # checked_siswa = ["Agus Ahmad", "Ajat Wahyudin"]
            for i in range(data_siswa.shape[0]):
                nm = data_siswa.loc[i,"Nama"]
                kls = data_siswa.loc[i,"Kelas"]
                new_kls = update_kelas(nm, kls, checked_siswa)
                data_siswa.loc[i,"Kelas"] = new_kls
            data_siswa.to_excel("./data/data_siswa_updated.xlsx", index=None)
            saveDataToDB("./data/data_siswa_updated.xlsx", "/data_siswa.xlsx")
            ## with open("./data/data_siswa_updated.xlsx", 'rb') as f:    # untuk reset data
            ##     dbx.files_upload(f.read(), "/data_siswa.xlsx", mode=dropbox.files.WriteMode.overwrite)
            return redirect(url_for("role"))

        if save_komentar=="1":
            # update Submit Komentar
            komentar = request.form['komentar']     # "Progres siswa sudah baik" # obtained from text box
            form_komentar.loc[(nisn_siswa), "Komentar"] = komentar
            # save komentar dataframe
            form_komentar.reset_index(inplace=True)
            form_komentar.to_excel("./nilai/Komentar_{}.xlsx".format(kelas), index=None)
            saveDataToDB("./nilai/Komentar_{}.xlsx".format(kelas), "/nilai/{}/{}/Komentar_{}.xlsx".format(folder_name_1, eval_type, kelas))
            ## with open("./nilai/Komentar_{}.xlsx".format(kelas), 'rb') as f:
            ##     dbx.files_upload(f.read(), "/nilai/{}/{}/Komentar_{}.xlsx".format(folder_name_1, eval_type, kelas), mode=dropbox.files.WriteMode.overwrite)


        if nisn_siswa != None and save_komentar!="1":
            # present nilai siswa
            nilai_siswa_sel = form_nilai.loc[[nisn_siswa]]
            komentar_siswa_sel = form_komentar.loc[nisn_siswa, "Komentar"]
            try:
                char_count = len(komentar_siswa_sel)
            except:
                char_count = 0

            # return str(len(komentar_siswa_sel))
            if pd.isna(komentar_siswa_sel):
                komentar_siswa_sel = ""
            nilai_sel = []
            # nama_mapel = []
            # # nilai_siswa_sel.to_excel("outs.xlsx")
            # # list mapel
            # # return str((nilai_siswa_sel.shape))
            # for i in range(nilai_siswa_sel.shape[1]):
            #     if i > 2 and i < nilai_siswa_sel.shape[1]-6 and i%3==0 :
            #         nama_mapel.append(form_nilai.columns[i].replace("_Keterampilan",""))
            for i in range(nilai_siswa_sel.shape[1]):
                if i > 1 and i < nilai_siswa_sel.shape[1]-1 :
                    if pd.isna(nilai_siswa_sel.iloc[0,i]):
                        nilai_sel.append("-")
                    else:
                        try:
                            nilai_sel.append("{0:.2f}".format(nilai_siswa_sel.iloc[0,i]))
                        except:
                            nilai_sel.append("{}".format(nilai_siswa_sel.iloc[0,i]))
                else:
                    nilai_sel.append(nilai_siswa_sel.iloc[0,i])
                # endloop 
            guru_mapel = loadDataInPd("/nilai/{}/guru_mapel.xlsx".format(folder_name_1), "/guru_mapel.xlsx")
            nama_mapel = guru_mapel["Mata Pelajaran"].values.tolist()
            # return str(nama_mapel)
            jlh_mapel = int((len(nama_mapel)))
            # print(nilai_sel)
            # nilai_siswa_sel.fillna("-", inplace=True)
            # nilai_sel = nilai_siswa_sel.values.tolist()[0]
            nilai_sel.insert(0,0)
            nilai_sel[2] = "{0:.2f}".format(nilai_sel[2])
            # print(tmp2)
            # tmp1.append(nilai_siswa_sel.values.tolist())
            # print(tmp)
            html= render_template("nilai_wali.html",nilai_sel=nilai_sel, tahun_ajaran=tahun_ajaran, cetak="0",lihat=lihat, nisn_siswa=nisn_siswa,
                                        pelajaran=pelajaran, kelas=kelas, semester=semester, nama_mapel=nama_mapel, komentar_siswa_sel=komentar_siswa_sel,char_count=char_count,
                                        eval_type=eval_type, nama_guru=session['nama_user'],jlh_mapel=jlh_mapel,jlh_nilai_sel=len(nilai_sel))

        if cetak_semua == "1":
            # Cetak Semua >> iterating to print the data of all student (./nilai/Name_NISN_Raport.pdf)
            # merging pdf file
            pdfs = []
            for siswa in range(len(list_siswa)):
                nilai_siswa_sel = form_nilai.loc[[list_nisn[siswa]]]
                komentar_siswa_sel = form_komentar.loc[list_nisn[siswa], "Komentar"]
                try:
                    char_count = len(komentar_siswa_sel)
                except:
                    char_count = 0
                if pd.isna(komentar_siswa_sel):
                    komentar_siswa_sel = ""
                nilai_sel = []
                # nama_mapel = []
                # for i in range(nilai_siswa_sel.shape[1]):
                #     if i > 2 and i < nilai_siswa_sel.shape[1]-6 and i%3==0 :
                #         nama_mapel.append(form_nilai.columns[i].replace("_Keterampilan",""))
                for i in range(nilai_siswa_sel.shape[1]):
                    if i > 1 and i < nilai_siswa_sel.shape[1]-1 :
                        if pd.isna(nilai_siswa_sel.iloc[0,i]):
                            nilai_sel.append("-")
                        else:
                            try:
                                nilai_sel.append("{0:.2f}".format(nilai_siswa_sel.iloc[0,i]))
                            except:
                                nilai_sel.append("{}".format(nilai_siswa_sel.iloc[0,i]))
                    else:
                        nilai_sel.append(nilai_siswa_sel.iloc[0,i])
                nilai_sel.insert(0,0)
                nilai_sel[2] = "{0:.2f}".format(nilai_sel[2])
                # print(nilai_sel[2])
                guru_mapel = loadDataInPd("/nilai/{}/guru_mapel.xlsx".format(folder_name_1), "/guru_mapel.xlsx")
                nama_mapel = guru_mapel["Mata Pelajaran"].values.tolist()
                # return str(nama_mapel)
                jlh_mapel = int((len(nama_mapel)))

                html= render_template("nilai_wali.html",nilai_sel=nilai_sel, tahun_ajaran=tahun_ajaran, cetak="0",lihat=lihat, nisn_siswa=list_nisn[siswa],
                                            pelajaran=pelajaran, kelas=kelas, semester=semester, nama_mapel=nama_mapel, komentar_siswa_sel=komentar_siswa_sel,char_count=char_count,
                                            eval_type=eval_type, nama_guru=session['nama_user'],jlh_mapel=jlh_mapel,jlh_nilai_sel=len(nilai_sel))

                filename_pdf = "./nilai/{}_{}_Raport.pdf".format(list_siswa[siswa], list_nisn[siswa])
                pdf = createPdf(html, filename_pdf)
                pdfs.append(filename_pdf)

            merger = PdfFileMerger()
            for pdf in pdfs:
                merger.append(pdf,import_bookmarks=False)
            merger.write("nilai/Raport_{}.pdf".format(kelas))   # Saved in folder nilai
            merger.close()
            filenames = "Raport_{}.pdf".format(kelas)
            return redirect(url_for('download_from_directory', source='nilai',filename=filenames))
        elif cetak_semua =="0":
            if lihat == "0":
                # Cetak >> print the data of selected student (./nilai/Name_NISN_Raport.pdf)
                # print pdf
                filename_pdf = "{}_{}_Raport.pdf".format(nilai_sel[1], nisn_siswa)
                # filename_pdf = "Raport_{}_{}_{}.pdf".format(pelajaran,kelas,nilai_sel[1])
                pdf = createPdf(html)
                response = createResponse(pdf, filename_pdf)
                return response
            elif lihat =="1":
                return html
        return render_template("rekap_wali.html",jlh_list=len(list_nisn), list_nisn=list_nisn, list_siswa=list_siswa, eval_type=eval_type, pelajaran=pelajaran, kelas=kelas)
    else:
        return redirect(url_for("login"))

# ================           ============================

@app.route("/ganti-pass",methods=["GET", "POST"])
def ganti_password():
    if "user" in session:
        if request.method=="POST":

            # ganti password
            # load data guru
            data_guru = loadDataInPd("/data_guru.xlsx")
            id_guru = session["user"]  # 106 # <from input>
            # input
            passwd_lama = request.form["pass_lama"] #"smpbinarilmu"
            passwd_baru = request.form["pass_baru"] #"mypassword"
            passwd_retype = request.form["pass_baru_re"] #"mypassword"

            data_guru.set_index("ID Guru", inplace=True)

            passwd_prev = data_guru.loc[id_guru, "Password"]

            if passwd_lama != passwd_prev:
                print("Password sebelumnya salah")
            elif passwd_baru != passwd_retype:
                print("Password baru tidak konsisten")
            else:
                # update password lama
                data_guru.loc[id_guru, "Password"] = passwd_baru
            data_guru.reset_index(inplace=True)
            data_guru.to_excel("./data/data_guru.xlsx", index=None)
            saveDataToDB("./data/data_guru.xlsx", "/data_guru.xlsx")
            ## with open("./data/data_guru.xlsx", 'rb') as f:
            ##     dbx.files_upload(f.read(), "/data_guru.xlsx", mode=dropbox.files.WriteMode.overwrite)
            return redirect(url_for("role"))
        return render_template("ganti_pass.html", id=session["user"],nama=session["nama_user"])
    else:
        return redirect(url_for("login"))

@app.route('/input/<string:pelajaran>/<string:kelas>')
def download_template_nilai(pelajaran, kelas):
    filenames = "form_nilai_{}_{}.xlsx".format(pelajaran, kelas)
    return send_from_directory(PATH_TMP, path=filenames, as_attachment=True)

@app.route('/download/<string:source>/<string:filename>')
def download_from_directory(source, filename):
    if source == "tmp":
        return send_from_directory(PATH_TMP, path=filename, as_attachment=True)
    elif source == "data":
        return send_from_directory(PATH_DATA, path=filename, as_attachment=True)
    elif source == "nilai":
        return send_from_directory(PATH_NILAI, path=filename, as_attachment=True)

@app.route("/clear")
def clearSession():
    session.pop('tingkat',None)
    session.pop('user',None)
    session.pop('nama_user',None)
    return redirect(url_for('login'))

# ====================== Non Route =======================================
def getAspekMateriList(aspek_materi):
    aspek_materi = aspek_materi.split(";")
    tmp = aspek_materi[:]
    aspek_materi = []
    # print(tmp)
    for materi in tmp:
        materi = materi.split(" ")
        if "" in materi:
            materi.remove("")
        materi = ' '.join(materi)
        if materi.isspace() == False:
            aspek_materi.append(materi)
    # print(aspek_materi)
    return aspek_materi
    
def save_template_form_nilai(pelajaran, kelas, aspek_materi, eval_type):
    # 5. Unduh dan unggah form nilai
    tahun_ajaran, semester, folder_name_1 = check_period()
    # load data siswa
    data_siswa = loadDataInPd("/nilai/{}/data_siswa.xlsx".format(folder_name_1), "/data_siswa.xlsx")
    # define aspek penilaian
    aspek_lain = ["Spiritual_Predikat", "Sosial_Predikat"]
    aspek_materi = getAspekMateriList(aspek_materi)
    # cek validitas

    # Unduh form nilai
    # create dataframe of form penilaian
    form_nilai = data_siswa[data_siswa["Kelas"] == kelas][["NISN", "Nama"]]
    for aspek in aspek_lain:
        form_nilai[aspek] = ""
    for aspek in aspek_materi:
        for i in ["Pengetahuan", "Keterampilan"]:
            form_nilai["{}_{}".format(aspek, i)] = ""
    form_nilai["{}_Pengetahuan".format(eval_type)] = ""
    form_nilai["Komentar"] = ""

    form_nilai.to_excel("tmp/form_nilai_{}_{}.xlsx".format(pelajaran, kelas), index=None)

def unggah_form_nilai(file, eval_type, pelajaran, kelas, aspek_materi):
    # grab upload excel name
    excel_name = file.filename
    file_destination = "/".join([PATH_NILAI,excel_name])
    file.save(file_destination)

    # Unggah form nilai
    check_folder(eval_type)
    # uploaded file is saved in nilai
    form_nilai = pd.read_excel(file_destination)
    # update nilai Pengetahuan
    # form_nilai["Pengetahuan"] = 0
    # for i in range(form_nilai.shape[0]):
    #     tmp = np.average(form_nilai.iloc[:,6:-4].values)
    #     form_nilai.loc[i,"Pengetahuan"] = tmp
    # form_nilai.to_excel(file_destination, index=None)
    # load form nilai
    # update komentar
    form_komentar = form_nilai.iloc[:,[0,1,-1]]
    tahun_ajaran, semester, folder_name_1 = check_period()
    form_komentar.to_excel("./nilai/Komentar_{}_{}.xlsx".format(pelajaran, kelas), index=None)
    saveDataToDB("./nilai/Komentar_{}_{}.xlsx", "/nilai/{}/{}/Komentar_{}_{}.xlsx".format(folder_name_1, eval_type, pelajaran, kelas))
    ## with open("./nilai/Komentar_{}_{}.xlsx".format(pelajaran, kelas), 'rb') as f:
    ##     dbx.files_upload(f.read(), "/nilai/{}/{}/Komentar_{}_{}.xlsx".format(folder_name_1, eval_type, pelajaran, kelas), mode=dropbox.files.WriteMode.overwrite)
    
    form_nilai = form_nilai.iloc[:,:-1]
    col_title = form_nilai.columns.tolist()
    list_pengetahuan = []; list_keterampilan = [];  list_aspek = []
    for title in col_title:
        tmp = title.split("_")
        if len(tmp) > 1:
            if tmp[1] == "Pengetahuan":
                list_pengetahuan.append(title)
            elif tmp[1] == "Keterampilan":
                list_aspek.append(tmp[0])
                list_keterampilan.append(title)
    
    # validasi
    # load NIS sesuai kelas
    data_siswa = loadDataInPd("/nilai/{}/data_siswa.xlsx".format(folder_name_1))
    data_siswa_kelas = data_siswa[data_siswa["Kelas"] == kelas]
    list_nisn = data_siswa_kelas["NISN"].values.tolist()
    list_nisn_form = form_nilai["NISN"].values.tolist()

    flag = 0
    if len(list_nisn) != len(list_nisn_form):
        # print("jumlah siswa tidak sama")
        flag = 1
    
    if len(aspek_materi) != len(list_aspek):
        # print("jumlah aspek materi tidak sama")
        flag = 1
    
    if flag == 0:
        for i in range(len(list_nisn)):
            if list_nisn[i] != list_nisn_form[i]:
                # print("nisn tidak sama")
                flag = 1
                break
        
        for i in range(len(aspek_materi)):
            if aspek_materi[i] != list_aspek[i]:
                # print("aspek materi tidak sama")
                flag = 1
                break

    if flag == 1:
        msg = "File yang diunggah tidak valid, silahkan unggah ulang"
        return flag, msg

    nilai_pengetahuan = form_nilai[list_pengetahuan]
    nilai_pengetahuan["Harian_Avg"] =  nilai_pengetahuan.iloc[:,:-1].mean(axis=1)
    tmp = nilai_pengetahuan.iloc[:,-1]*.75 + nilai_pengetahuan.iloc[:,-2]*.25
    form_nilai["Nilai Akhir Pengetahuan"] = tmp
    nilai_keterampilan = form_nilai[list_keterampilan]
    nilai_keterampilan["Harian_Avg"] =  nilai_keterampilan.mean(axis=1)
    form_nilai["Nilai Akhir Keterampilan"] = nilai_keterampilan["Harian_Avg"]
    form_nilai["Spiritual_Deskripsi"] = form_nilai["Spiritual_Predikat"].apply(lambda x: deskripsi(x, type="spiritual"))
    form_nilai["Sosial_Deskripsi"] = form_nilai["Spiritual_Predikat"].apply(lambda x: deskripsi(x, type="sosial"))

    col_title = form_nilai.columns.tolist()
    form_nilai_new = form_nilai[col_title[:3]]
    form_nilai_new["Spiritual_Deskripsi"] = form_nilai["Spiritual_Predikat"].apply(lambda x: deskripsi(x, type="spiritual"))
    form_nilai_new[col_title[3]] = form_nilai[col_title[3]]
    form_nilai_new["Sosial_Deskripsi"] = form_nilai["Spiritual_Predikat"].apply(lambda x: deskripsi(x, type="sosial"))
    for title in col_title[4:]:
        form_nilai_new[title] = form_nilai[title]

    form_nilai = form_nilai_new

    form_nilai.to_excel(file_destination, index=None)

    tahun_ajaran, semester, folder_name_1 = check_period()
    saveDataToDB(file_destination, "/nilai/{}/{}/form_nilai_{}_{}.xlsx".format(folder_name_1, eval_type, pelajaran, kelas))
    ## with open(file_destination, 'rb') as f:
    ##     dbx.files_upload(f.read(), "/nilai/{}/{}/form_nilai_{}_{}.xlsx".format(folder_name_1, eval_type, pelajaran, kelas), mode=dropbox.files.WriteMode.overwrite)
    
    # update input status
    # load status file
    status_nilai = loadDataInPd("/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type))
    status_nilai.set_index("Mata Pelajaran", inplace=True)
    status_nilai.loc[pelajaran, kelas] = 1
    status_nilai.to_excel("./tmp/status_nilai.xlsx")
    saveDataToDB("./tmp/status_nilai.xlsx", "/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type))
    ## with open("./tmp/status_nilai.xlsx", 'rb') as f:
    ##     dbx.files_upload(f.read(), "/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type), mode=dropbox.files.WriteMode.overwrite)
    return flag, None

def deskripsi(x, type="spiritual"):
    if type == "spiritual":
        if x == 1:
            predikat = "Kurang. Kurang menerima dan tidak konsisten menjalankan, menghargai, menghayati dan mengamalkan nilai agama."
        elif x == 2:
            predikat = "Cukup. Cukup menerima dan belum konsisten menjalankan, menghargai, menghayati dan mengamalkan nilai agama."
        elif x == 3:
            predikat = "Baik. Dapat menerima dan mulai konsisten menjalankan, menghargai, menghayati dan mengamalkan nilai agama."
        else:
            predikat = "Sangat baik. Sangat menerima dan sudah konsisten menjalankan, menghargai, menghayati dan mengamalkan nilai agama."
    else:
        if x == 1:
            predikat = "Kurang. Kurang menerima dan tidak konsisten menjalankan, menghargai dan menghayati perilaku jujur, disiplin, tanggung jawab, peduli, toleransi, gotong royong, santun, percaya diri dalam berinteraksi secara efektif dengan lingkungan sosial dan alam dalam jangkauan pergaulan dan keberadaannya."
        elif x == 2:
            predikat = "Cukup. Cukup menerima dan belum konsisten menjalankan, menghargai dan menghayati perilaku jujur, disiplin, tanggung jawab, peduli, toleransi, gotong royong, santun, percaya diri dalam berinteraksi secara efektif dengan lingkungan sosial dan alam dalam jangkauan pergaulan dan keberadaannya."
        elif x == 3:
            predikat = "Baik. Dapat menerima dan mulai konsisten menjalankan, menghargai dan menghayati perilaku jujur, disiplin, tanggung jawab, peduli, toleransi, gotong royong, santun, percaya diri dalam berinteraksi secara efektif dengan lingkungan sosial dan alam dalam jangkauan pergaulan dan keberadaannya."
        else:
            predikat = "Sangat baik. Sangat menerima dan sudah konsisten menjalankan, menghargai dan menghayati perilaku jujur, disiplin, tanggung jawab, peduli, toleransi, gotong royong, santun, percaya diri dalam berinteraksi secara efektif dengan lingkungan sosial dan alam dalam jangkauan pergaulan dan keberadaannya."
    return predikat

def createPdf(html, filename_pdf=False):
    css = ["static/css/bootstrap.min.css","static/style.css"]
    ## ============   uncomment config yang dipilih  =================
        # config for heroku :
    config = pdfkit.configuration(wkhtmltopdf='/app/bin/wkhtmltopdf')
        # config for local ver 1 :
    # config = pdfkit.configuration(wkhtmltopdf='./bin/wkhtmltopdf')
    # config = pdfkit.configuration(wkhtmltopdf="/usr/local/bin/wkhtmltopdf")
        # config for local Windows :
    # path_wkhtmltopdf = "C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe"
    # config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

    pdf = pdfkit.from_string(html, filename_pdf,configuration=config, css=css)
    # config for local pc
    # pdf = pdfkit.from_string(html, filename_pdf, css=css)
    return pdf

def createResponse(pdf,filename_pdf):
    headers_filename = "attachment; filename="+filename_pdf
    response = make_response(pdf)
    response.headers["Content-Type"] = "application/pdf"
    response.headers["Content-Disposition"] = headers_filename
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    response.headers['Cache-Control'] = 'public, max-age=0'
    return response

# dev only
@app.route('/get-data-dev/<string:filename>')
def getDataDev(filename):
    # file_stream=stream_dropbox_file("/data_guru.xlsx")
    # data = pd.read_excel(file_stream)
    dbx.files_download_to_file("./dev/{}".format(filename), "{}/{}".format(check_jenjang(), filename))
    # data = pd.read_excel("./tmp/data_guru.xlsx")
    return filename

@app.route('/del_foto/<string:filename>/<string:tingkat>')
def delFoto(filename, tingkat):
    if tingkat == "SMA":
        jenjang = "/SMA"
    elif tingkat == "SMP":
        jenjang = ""
    else:
        return filename

    try:
        dbx.files_delete("{}/dokumen/pas_foto/{}".format(jenjang, filename))
    except:
        pass
    try:
        dbx.files_delete("{}/dokumen/ijazah_sd/{}".format(jenjang, filename))
    except:
        pass
    try:
        dbx.files_delete("{}/dokumen/kartu_keluarga/{}".format(jenjang, filename))
    except:
        pass
    try:
        dbx.files_delete("{}/dokumen/akta_kelahiran/{}".format(jenjang, filename))
    except:
        pass
    return filename

@app.after_request
def add_header(r):
    """
    Add headers to both force latest IE rendering engine or Chrome Frame,
    and also to cache the rendered page for 10 minutes.
    """
    r.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    r.headers["Pragma"] = "no-cache"
    r.headers["Expires"] = "0"
    r.headers['Cache-Control'] = 'public, max-age=0'
    return r

if __name__ == "__main__":
    app.run(debug=True)