import re
from flask import Flask, redirect, url_for, render_template, request, send_from_directory, make_response, session, current_app
import os
from flask.helpers import send_file
from numpy.lib.npyio import save
import pandas as pd
import numpy as np
import csv
import pdfkit
import datetime as dtm
from datetime import date,datetime

import PyPDF2
import dropbox
from contextlib import closing
from io import BytesIO
from PyPDF2 import PdfFileMerger
from six import ensure_text

app = Flask(__name__)
app.secret_key = "penilaian-sekolah-IKN"
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0

APP_ROOT = os.path.dirname(os.path.abspath(__file__))
path_nilai = os.path.join(APP_ROOT, 'nilai/')
path_nilai = os.path.join(APP_ROOT, 'tmp/')

def stream_dropbox_file(path):
    _,res = dbx.files_download(path)
    with closing(res) as result:
        byte_data=result.content
        return BytesIO(byte_data)

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

def check_folder(eval_type):
    tahun_ajaran, semester, folder_name_1 = check_period()
    # create folder of year_semester if not exist
    try:
        dbx.files_get_metadata("/nilai/{}".format(folder_name_1))
    except:
        dbx.files_create_folder("/nilai/{}".format(folder_name_1))
    # create folder PTS/PAS if not exist
    try:
        dbx.files_get_metadata("/nilai/{}/{}".format(folder_name_1, eval_type))
    except:
        dbx.files_create_folder("/nilai/{}/{}".format(folder_name_1, eval_type))

# update kelas
def update_kelas(nm, kls, list_nm):
    if nm in list_nm:
        if kls == "VII":
            new_kls = "VIII"
        elif kls == "VIII":
            new_kls = "IX"
        else:
            new_kls = "Alumni"
    else:
        new_kls = kls
    return new_kls

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

# 1. Halaman login
# login to dropbox
token = "UUqaOMObW8sAAAAAAAAAAcvkmOaxYzJs7ZRhjCRDqVqvKuP-8Gd1W0n7i6CjhKNK"
dbx = dropbox.Dropbox(token)

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
                # verifikasi akun
                # misal input id_guru di-assign sebagai variable "id_guru"
                try:
                    id_guru = int(request.form['id_guru'])
                except:
                    return render_template("login.html",message="pwdSalah")
                # misal input password di-assign sebagai variable "input_passwd"
                input_passwd = request.form['password']
                if id_guru == 100 and input_passwd=="adminsmpbinarilmu":
                    session["nama_user"] = "Admin"
                    session["user"] = id_guru
                    return redirect(url_for("admin"))

                # load data guru
                file_stream=stream_dropbox_file("/data_guru.xlsx")
                data_guru = pd.read_excel(file_stream)
                # dbx.files_download_to_file("./tmp/data_guru.xlsx", "/data_guru.xlsx")
                # data_guru = pd.read_excel("./tmp/data_guru.xlsx")

                list_ID_guru = data_guru.loc[:, "ID Guru"].tolist()
                if id_guru not in list_ID_guru:
                    # print("ID guru tidak terdaftar")
                    return render_template("login.html",message="error")
                else:
                    passwd = data_guru[data_guru["ID Guru"]==id_guru]["Password"].values[0]
                    if input_passwd == passwd:
                        print("OK") # masuk ke halaman selanjutnya (role)
                        nama_guru = data_guru[data_guru["ID Guru"]==id_guru]["Nama"].values[0]
                        session["nama_user"] = nama_guru
                        session["user"] = id_guru
                        return redirect(url_for("role"))
                    else:
                        # tampilkan kalimat "Password salah"
                        # print("Password salah")
                        return render_template("login.html",message="pwdSalah")

        return render_template("login.html")

@app.route("/admin",methods=["GET", "POST"])
def admin():
    if session["user"] == 100:
        if request.method=="POST":
            pilihan = request.form['pilihan']
            return redirect(url_for(pilihan))
        return render_template("admin.html")
    else:
        return redirect(url_for("login"))

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
                file_stream=stream_dropbox_file("/data_siswa.xlsx")
                data_siswa = pd.read_excel(file_stream)
                data_kelas = data_siswa[data_siswa["Kelas"] == kelas]
                list_nisn = data_kelas["NISN"].tolist()
                list_siswa = data_kelas["Nama"].tolist()
                list_kelas_siswa = data_kelas["Kelas"].tolist()

                # return str(list_kelas)
                return render_template("hal_kelas.html",kelas=kelas,list_kelas=list_kelas, list_nisn=list_nisn, jlh_list=len(list_nisn),
                                        list_siswa=list_siswa, list_kelas_siswa=list_kelas_siswa)
        # 9 Halaman Data Siswa
        # load data siswa
        file_stream=stream_dropbox_file("/data_siswa.xlsx")
        data_siswa = pd.read_excel(file_stream)
        list_kelas = data_siswa["Kelas"].unique()
        list_kelas = list(list_kelas)
        # for kelas in list_kelas:
        #     print("Kelas {}".format(kelas))
        return render_template("data_siswa.html",list_kelas=list_kelas,jlh_kelas=len(list_kelas))
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

            file_stream=stream_dropbox_file("/data_siswa.xlsx")
            data_siswa = pd.read_excel(file_stream)
            # nisn = 72100585
            # nama = "Agus Ahmad"  # <from input>
            # kelas = "VII"    # <from input>
            data_siswa.set_index("NISN", inplace=True)

            if save_edit=="1":
                # return nisn_edit
                # return str(kelas_edit+" "+nama_edit)
                # Tombol Edit
                data_siswa.loc[nisn_edit, "Nama"] = nama_edit
                data_siswa.loc[nisn_edit, "Kelas"] = kelas_edit
                data_siswa.reset_index(inplace=True)


                data_siswa.to_excel("./data/data_siswa.xlsx", index=None)
                with open("./data/data_siswa.xlsx", 'rb') as f:
                    dbx.files_upload(f.read(), "/data_siswa.xlsx", mode=dropbox.files.WriteMode.overwrite)
                return redirect(url_for("data_siswa"))
            elif dropout =="1":
                # return "DO"
                # Tombol Dropout
                data_siswa.loc[nisn_edit, "Kelas"] = "Dropout"
                data_siswa.reset_index(inplace=True)

                data_siswa.to_excel("./data/data_siswa.xlsx", index=None)
                with open("./data/data_siswa.xlsx", 'rb') as f:
                    dbx.files_upload(f.read(), "/data_siswa.xlsx", mode=dropbox.files.WriteMode.overwrite)
                return redirect(url_for("data_siswa"))
        return render_template("edit_siswa.html",kelas=kelas,list_kelas=list_kelas, nisn_edit=nisn_edit,
                                        nama_edit=nama_edit, kelas_edit=kelas_edit)

    else:
        return redirect(url_for("login"))

@app.route("/admin/data-siswa/tambah",methods=["GET", "POST"])
def tambah_siswa():
    if session["user"] == 100:
        if request.method == "POST":
            # 12 Halaman Tambah Siswa
            # Unduh
            # sent ./data/data_siswa_template.xlsx to user

            kelas = request.form["kelas"]
            try:
                tambah = request.form["unggah"]
                if tambah =="1":
                    file = request.files["file"]
                    excel_name = file.filename
                    file_destination = "/".join([path_nilai,excel_name])
                    file.save(file_destination)

                    # Unggah
                    # uploaded file is located at ./tmp/data_siswa_template.xlsx to
                    new_data = pd.read_excel(file_destination)
                    file_stream=stream_dropbox_file("/data_siswa.xlsx")
                    data_siswa = pd.read_excel(file_stream)
                    data_siswa = pd.concat([data_siswa, new_data], axis=0)

                    data_siswa.to_excel("./data/data_siswa.xlsx", index=None)
                    # return redirect(url_for("data_siswa"))
                    # return "gg"
                    with open("./data/data_siswa.xlsx", 'rb') as f:
                    # with open("./data/data_siswas_hasil.xlsx", 'rb') as f:    # untuk reset data
                        dbx.files_upload(f.read(), "/data_siswa.xlsx", mode=dropbox.files.WriteMode.overwrite)
                    return redirect(url_for("data_siswa"))
            except:
                return render_template("tambah_siswa.html")
    else:
        return redirect(url_for("login"))

@app.route("/admin/data-guru",methods=["GET", "POST"])
def data_guru():
    if session["user"] == 100:
        if request.method=="POST":
            pilihan = request.form['pilihan']
            return pilihan
            return render_template("role_mapel.html", nama_mapel=nama_mapel, kelas=kelas)
        # 13 Halaman Data Guru
        # load data guru
        file_stream=stream_dropbox_file("/data_guru.xlsx")
        data_guru = pd.read_excel(file_stream)
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

            id_sel = int(request.form["id_sel"])
            nama_sel = request.form["nama_sel"]
            pass_sel = request.form["pass_sel"]

            if save_edit=="1":
                # return str(nama_sel+" "+pass_sel)
                # 14 Halaman Edit Data Guru
                # load data guru
                file_stream=stream_dropbox_file("/data_guru.xlsx")
                data_guru = pd.read_excel(file_stream)

                # id_guru = 101
                # nama = "Wahyu Mediana"  # <from input>
                # password = "smpbinarilmu" # <from input>

                data_guru.set_index("ID Guru", inplace=True)
                data_guru.loc[id_sel, "Nama"] = nama_sel
                data_guru.loc[id_sel, "Password"] = pass_sel
                data_guru.reset_index(inplace=True)

                data_guru.to_excel("./data/data_guru.xlsx", index=None)
                with open("./data/data_guru.xlsx", 'rb') as f:
                    dbx.files_upload(f.read(), "/data_guru.xlsx", mode=dropbox.files.WriteMode.overwrite)
                return redirect(url_for("data_guru"))
            return render_template("edit_guru.html", id_sel=id_sel,
                                        nama_sel=nama_sel, pass_sel=pass_sel)

    else:
        return redirect(url_for("login"))

@app.route("/admin/data-guru/tambah",methods=["GET", "POST"])
def tambah_guru():
    if session["user"] == 100:
        if request.method == "POST":
            pass
            try:
                tambah = request.form["tambah"]
                name = request.form["nama"]

            except:
                tambah = "0"
            # 15 Halaman Tambah Data Guru
            # load data guru
            file_stream=stream_dropbox_file("/data_guru.xlsx")
            data_guru = pd.read_excel(file_stream)
            id_now = data_guru["ID Guru"].tolist()[-1] + 1
            # name = "Isman Kurniawan"   # <from input>
            if tambah == "1":
                data_guru = data_guru.append({'ID Guru': id_now, 'Nama': name, 'Password': 'smpbinarilmu'}, ignore_index=True)
                data_guru.to_excel("./data/data_guru.xlsx", index=None)
                with open("./data/data_guru.xlsx", 'rb') as f:
                    dbx.files_upload(f.read(), "/data_guru.xlsx", mode=dropbox.files.WriteMode.overwrite)
                return redirect(url_for("data_guru"))
            return render_template("tambah_guru.html",id_now=id_now)
    else:
        return redirect(url_for("login"))


@app.route("/admin/plotting-pengajaran",methods=["GET", "POST"])
def plot_pengajaran():
    if session["user"] == 100:

        # 16 Halaman Plotting Pengajaran
        # load guru mapel
        file_stream=stream_dropbox_file("/guru_mapel.xlsx")
        guru_mapel = pd.read_excel(file_stream)
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
            newData = pd.DataFrame(list(zip(list_mapel, list_VII_baru,list_VIII_baru,list_IX_baru)),
               columns =['Mata Pelajaran', 'VII', 'VIII', 'IX'])

            newData.to_excel("./data/guru_mapel.xlsx", index=None)
            # update guru_mapel.xlsx according to the input
            # guru_mapel.to_excel("./data/guru_mapel.xlsx", index=None)
            with open("./data/guru_mapel.xlsx", 'rb') as f:
                dbx.files_upload(f.read(), "/guru_mapel.xlsx", mode=dropbox.files.WriteMode.overwrite)
            return redirect(url_for("plot_pengajaran"))
        # get list guru
        # load data guru
        file_stream=stream_dropbox_file("/data_guru.xlsx")
        data_guru = pd.read_excel(file_stream)
        list_guru = data_guru["Nama"].tolist()



        return render_template("plot_pengajaran.html",list_mapel=list_mapel,jlh_list=len(list_mapel),list_guru=list_guru,
                                   list_kelasVII=list_kelasVII, list_kelasVIII=list_kelasVIII,list_kelasIX=list_kelasIX,jlh_guru=len(list_guru))
    else:
        return redirect(url_for("login"))

@app.route("/admin/plotting-wali",methods=["GET", "POST"])
def plot_wali():
    if session["user"] == 100:

        # 16 Halaman Plotting Pengajaran
        # load guru mapel
        file_stream=stream_dropbox_file("/guru_wali_kelas.xlsx")
        guru_wali_kelas = pd.read_excel(file_stream)
        # buat cek data
        guru_wali_kelas.to_excel("./data/guru_wali_kelas.xlsx", index=None)
        # return "test plot wali"
        list_wali = guru_wali_kelas["Wali Kelas"].tolist()
        list_kelas = guru_wali_kelas["Kelas"].tolist()
        # get list guru
        # load data guru
        file_stream=stream_dropbox_file("/data_guru.xlsx")
        data_guru = pd.read_excel(file_stream)
        list_guru = data_guru["Nama"].tolist()

        if request.method=="POST":
            list_wali_baru = request.form.getlist('list_wali[]')
            newData = pd.DataFrame(list(zip(list_kelas, list_wali_baru)),
               columns =['Kelas', 'Wali Kelas'])

            newData.to_excel("./data/guru_wali_kelas_baru.xlsx", index=None)
            # update guru_mapel.xlsx according to the input
            ### guru_mapel.to_excel("./data/guru_mapel.xlsx", index=None)
            with open("./data/guru_wali_kelas_baru.xlsx", 'rb') as f:
                dbx.files_upload(f.read(), "/guru_wali_kelas.xlsx", mode=dropbox.files.WriteMode.overwrite)
            return redirect(url_for("plot_wali"))

        return render_template("plot_wali.html",jlh_list=len(list_kelas),list_guru=list_guru,
                                   jlh_guru=len(list_guru),list_kelas=list_kelas,list_wali=list_wali)
    else:
        return redirect(url_for("login"))

@app.route("/ganti-pass",methods=["GET", "POST"])
def ganti_password():
    if "user" in session:
        if request.method=="POST":

            # ganti password
            # load data guru
            file_stream=stream_dropbox_file("/data_guru.xlsx")
            data_guru = pd.read_excel(file_stream)
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
            with open("./data/data_guru.xlsx", 'rb') as f:
                dbx.files_upload(f.read(), "/data_guru.xlsx", mode=dropbox.files.WriteMode.overwrite)
            return redirect(url_for("role"))
        return render_template("ganti_pass.html", id=session["user"],nama=session["nama_user"])
    else:
        return redirect(url_for("login"))
@app.route("/role",methods=["GET", "POST"])
def role():
    if "user" in session:
        # 2. Pilih Role
        # load data guru
        file_stream=stream_dropbox_file("/data_guru.xlsx")
        data_guru = pd.read_excel(file_stream)
        # create dict ID guru: nama guru
        dict_guru = {}
        for i in range(data_guru.shape[0]):
            dict_guru[data_guru.loc[i,"ID Guru"]] = data_guru.loc[i,"Nama"]
        # load data guru mapel
        file_stream=stream_dropbox_file("/guru_mapel.xlsx")
        guru_mapel = pd.read_excel(file_stream)
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
        file_stream=stream_dropbox_file("/guru_wali_kelas.xlsx")
        guru_wali_kelas = pd.read_excel(file_stream)
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
            pelajaran = request.form['mapel']   #"Wali Kelas"
            kelas = request.form['kelas']       #"VII"
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
                    dbx.files_get_metadata("/nilai/{}/{}/Rekap_Nilai_{}".format(folder_name_1, eval_type, kelas))
                    generate = False
                except:
                    generate = True
            # generate rekap nilai
            if generate:
                # load list_mapel
                file_stream=stream_dropbox_file("/guru_mapel.xlsx")
                guru_mapel = pd.read_excel(file_stream)
                guru_mapel.set_index("Mata Pelajaran", inplace=True)
                # create list mapel and kelas
                mapel = guru_mapel.index.tolist()
                # create list of exist file in folder nilai/tahun_sem/PTS or PAS
                a = dbx.files_list_folder(path="/nilai/{}/{}".format(folder_name_1, eval_type, kelas))
                file_list = []
                for i in range(len(a.entries)):
                    file_name = a.entries[i].name
                    file_list.append(file_name.split(".")[0])
                # create file rekap nilai
                # load data siswa
                file_stream=stream_dropbox_file("/data_siswa.xlsx")
                data_siswa = pd.read_excel(file_stream)
                form_nilai = data_siswa[data_siswa["Kelas"] == kelas][["NISN", "Nama"]]
                for mpl in mapel:
                    for aspek in ["Sikap", "Keterampilan", "Pengetahuan"]:
                        form_nilai["{}_{}".format(mpl, aspek)] = ""
                    file_name = "form_nilai_{}_{}".format(mpl, kelas)
                    if file_name in file_list:
                        # load file name
                        file_stream=stream_dropbox_file("/nilai/{}/{}/{}.xlsx".format(folder_name_1, eval_type, file_name))
                        form_nilai_mapel = pd.read_excel(file_stream)
                        nilai_sikap = form_nilai_mapel["Spiritual_Predikat"] + form_nilai_mapel["Sosial_Predikat"]
                        tmp = round(nilai_sikap/2,0)
                        form_nilai["{}_Sikap".format(mpl)] = tmp
                        form_nilai["{}_Pengetahuan".format(mpl)] = form_nilai_mapel["Nilai Akhir Pengetahuan"].values
                        form_nilai["{}_Keterampilan".format(mpl)] = form_nilai_mapel["Nilai Akhir Keterampilan"].values
                aspek_dict = {}; pred_dict = {}
                for i in ["Sikap", "Pengetahuan", "Keterampilan"]:
                    aspek_dict[i] = []
                    pred_dict[i] = []
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

                form_nilai.to_excel("./nilai/Rekap_Nilai_{}.xlsx".format(kelas))
                with open("./nilai/Rekap_Nilai_{}.xlsx".format(kelas), 'rb') as f:
                    dbx.files_upload(f.read(), "/nilai/{}/{}/Rekap_Nilai_{}.xlsx".format(folder_name_1, eval_type, kelas), mode=dropbox.files.WriteMode.overwrite)
                # create file komentar
                # create komentar dataframe
                komentar = 0
                try:
                    dbx.files_get_metadata("/nilai/{}/{}/Komentar_{}.xlsx".format(folder_name_1, eval_type, kelas))
                except:
                    komentar = 1
                if komentar:
                    form_komentar = data_siswa[data_siswa["Kelas"] == kelas][["NISN", "Nama"]]
                    form_komentar["Komentar"] = ""
                    form_komentar.to_excel("./nilai/Komentar_{}.xlsx".format(kelas), index=None)
                    with open("./nilai/Komentar_{}.xlsx".format(kelas), 'rb') as f:
                        dbx.files_upload(f.read(), "/nilai/{}/{}/Komentar_{}.xlsx".format(folder_name_1, eval_type, kelas), mode=dropbox.files.WriteMode.overwrite)

                return redirect(url_for("wali_rekap",eval_type=eval_type, pelajaran=pelajaran, kelas=kelas))
        else:
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
            pelajaran = request.form['mapel']   #"Matematika"
            kelas = request.form['kelas']       #"VII"
            try:
                status = int(request.form['status'])
                tahun_ajaran, semester, folder_name_1 = check_period()
                if status == 0:
                    # load form nilai
                    # file_stream=stream_dropbox_file("/nilai/{}/{}/form_nilai_{}_{}.xlsx".format(folder_name_1, eval_type, pelajaran, kelas))
                    try:
                        dbx.files_delete("/nilai/{}/{}/form_nilai_{}_{}.xlsx".format(folder_name_1, eval_type, pelajaran, kelas))
                        # update input status
                        # load status file
                        file_stream=stream_dropbox_file("/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type))
                        status_nilai = pd.read_excel(file_stream)
                        status_nilai.set_index("Mata Pelajaran", inplace=True)
                        status_nilai.loc[pelajaran, kelas] = 0
                        status_nilai.to_excel("./tmp/status_nilai.xlsx")
                        with open("./tmp/status_nilai.xlsx", 'rb') as f:
                            dbx.files_upload(f.read(), "/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type), mode=dropbox.files.WriteMode.overwrite)
                    except:
                        print("gagal hapus data mapel")
            except:
                status = None
            # cek input
            # hasil = "hasil pilihan PTA PTS :" + eval_type + pelajaran + kelas
            # return hasil

            # check input status
            check_folder(eval_type)
            # copy template status nilai if not exist
            tahun_ajaran, semester, folder_name_1 = check_period()
            try:
                dbx.files_get_metadata("/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type))
            except:
                dbx.files_copy("/template_status_nilai.xlsx", "/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type))
            # load status nilai
            file_stream=stream_dropbox_file("/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type))
            status_nilai = pd.read_excel(file_stream)
            status_nilai.set_index("Mata Pelajaran", inplace=True)
            if status == None:
                status = status_nilai.loc[pelajaran, kelas]
            if status == 0:
                # go to input nilai
                return render_template("aspek_materi.html", eval_type=eval_type, pelajaran=pelajaran, kelas=kelas)
            else:
                # pass # go to rekap page
                return redirect(url_for("mapel_rekap", eval_type=eval_type, pelajaran=pelajaran, kelas=kelas))

        else:
            return redirect(url_for("role"))
    else:
        return redirect(url_for("login"))

@app.route("/statusnilai")
def statusnilai():
    eval_type = "PTS"
    tahun_ajaran, semester, folder_name_1 = check_period()
    try:
        dbx.files_get_metadata("/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type))
    except:
        dbx.files_copy("/template_status_nilai.xlsx", "/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type))
    # load status nilai
    file_stream=stream_dropbox_file("/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type))

    data = pd.read_excel(file_stream)

    data.to_excel("./data/statusnilai.xlsx", index=None)
    # with open("./data/statusnilai.xlsx", 'rb') as f:    # untuk reset data
    #     dbx.files_upload(f.read(), "/data_siswa.xlsx", mode=dropbox.files.WriteMode.overwrite)
    return "hehe"

@app.route("/input",methods=["GET", "POST"])
def menu_input():
    if "user" in session:
        if request.method=="POST":
            # 4. Input aspek materi
            aspek_materi = request.form['aspek_materi']
            eval_type = request.form['eval_type']   #"PTS"
            pelajaran = request.form['pelajaran']   #"Matematika"
            kelas = request.form['kelas']       #"VII"

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
                unggah_form_nilai(file,eval_type,pelajaran,kelas)
                return redirect(url_for("mapel_rekap", eval_type=eval_type, pelajaran=pelajaran, kelas=kelas, aspek_materi=aspek_materi))

            return render_template("unggah_nilai.html", aspek_materi=aspek_materi,
                                        eval_type=eval_type, pelajaran=pelajaran,
                                        kelas=kelas )
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
        file_stream=stream_dropbox_file("/nilai/{}/{}/Rekap_Nilai_{}.xlsx".format(folder_name_1, eval_type, kelas))
        form_nilai = pd.read_excel(file_stream)

        list_nisn = form_nilai["NISN"].tolist()
        list_siswa = form_nilai["Nama"].tolist()

        form_nilai.set_index("NISN", inplace=True)
        # load form komentar
        file_stream=stream_dropbox_file("/nilai/{}/{}/Komentar_{}.xlsx".format(folder_name_1, eval_type, kelas))
        form_komentar = pd.read_excel(file_stream)
        form_komentar.set_index("NISN", inplace=True)
        # show the list of siswa

        if update == "1":
            # return str(len(checked_siswa))
            # update kelas
            file_stream=stream_dropbox_file("/data_siswa.xlsx")
            data_siswa = pd.read_excel(file_stream)
            # checked_siswa = ["Agus Ahmad", "Ajat Wahyudin"]
            for i in range(data_siswa.shape[0]):
                nm = data_siswa.loc[i,"Nama"]
                kls = data_siswa.loc[i,"Kelas"]
                new_kls = update_kelas(nm, kls, checked_siswa)
                data_siswa.loc[i,"Kelas"] = new_kls
            # data_siswa.to_excel("./data/data_siswa_updated.xlsx", index=None)
            # with open("./data/data_siswa_updated.xlsx", 'rb') as f:    # untuk reset data
            #     dbx.files_upload(f.read(), "/data_siswa.xlsx", mode=dropbox.files.WriteMode.overwrite)
            return redirect(url_for("role"))

        if save_komentar=="1":
            # update Submit Komentar
            komentar = request.form['komentar']     # "Progres siswa sudah baik" # obtained from text box
            form_komentar.loc[(nisn_siswa), "Komentar"] = komentar
            # save komentar dataframe
            form_komentar.reset_index(inplace=True)
            form_komentar.to_excel("./nilai/Komentar_{}.xlsx".format(kelas), index=None)
            with open("./nilai/Komentar_{}.xlsx".format(kelas), 'rb') as f:
                dbx.files_upload(f.read(), "/nilai/{}/{}/Komentar_{}.xlsx".format(folder_name_1, eval_type, kelas), mode=dropbox.files.WriteMode.overwrite)


        if nisn_siswa != None and save_komentar!="1":
            # present nilai siswa
            nilai_siswa_sel = form_nilai.loc[[nisn_siswa]]
            komentar_siswa_sel = form_komentar.loc[nisn_siswa, "Komentar"]
            if pd.isna(komentar_siswa_sel):
                komentar_siswa_sel = ""
            nilai_sel = []
            nama_mapel = []
            # nilai_siswa_sel.to_excel("outs.xlsx")
            # return str((nilai_siswa_sel.shape))
            for i in range(nilai_siswa_sel.shape[1]):
                if i > 2 and i < nilai_siswa_sel.shape[1]-6 and i%3==0 :
                    nama_mapel.append(form_nilai.columns[i].replace("_Keterampilan",""))
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
            # return str(nama_mapel)
            jlh_mapel = int((len(nama_mapel)))
            html= render_template("nilai_wali.html",nilai_sel=nilai_sel, tahun_ajaran=tahun_ajaran, cetak="0",lihat=lihat, nisn_siswa=nisn_siswa,
                                        pelajaran=pelajaran, kelas=kelas, semester=semester, nama_mapel=nama_mapel, komentar_siswa_sel=komentar_siswa_sel,
                                        eval_type=eval_type, nama_guru=session['nama_user'],jlh_mapel=jlh_mapel,jlh_nilai_sel=len(nilai_sel))

        if cetak_semua == "1":
            # Cetak Semua >> iterating to print the data of all student (./nilai/Name_NISN_Raport.pdf)
            # merging pdf file
            pdfs = []
            for siswa in range(len(list_siswa)):
                nilai_siswa_sel = form_nilai.loc[[list_nisn[siswa]]]
                komentar_siswa_sel = form_komentar.loc[list_nisn[siswa], "Komentar"]
                if pd.isna(komentar_siswa_sel):
                    komentar_siswa_sel = ""
                nilai_sel = []
                nama_mapel = []
                for i in range(nilai_siswa_sel.shape[1]):
                    if i > 2 and i < nilai_siswa_sel.shape[1]-6 and i%3==0 :
                        nama_mapel.append(form_nilai.columns[i].replace("_Keterampilan",""))
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

                jlh_mapel = int((len(nama_mapel)))
                html= render_template("nilai_wali.html",nilai_sel=nilai_sel, tahun_ajaran=tahun_ajaran, cetak="0",lihat=lihat, nisn_siswa=list_nisn[siswa],
                                            pelajaran=pelajaran, kelas=kelas, semester=semester, nama_mapel=nama_mapel, komentar_siswa_sel=komentar_siswa_sel,
                                            eval_type=eval_type, nama_guru=session['nama_user'],jlh_mapel=jlh_mapel,jlh_nilai_sel=len(nilai_sel))

                filename_pdf = "./nilai/{}_{}_Raport.pdf".format(list_siswa[siswa], list_nisn[siswa])
                css = ["static/css/bootstrap.min.css","static/style.css"]
                ## uncomment config yang dipilih
                    # config for heroku :
                config = pdfkit.configuration(wkhtmltopdf='/app/bin/wkhtmltopdf')
                    # config for local :
                # config = pdfkit.configuration(wkhtmltopdf='./bin/wkhtmltopdf')
                # config = pdfkit.configuration(wkhtmltopdf="/usr/local/bin/wkhtmltopdf")
                    # config for local windows :
                # path_wkhtmltopdf = "C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe"
                # config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

                pdf = pdfkit.from_string(html, filename_pdf,configuration=config, css=css)
                # config for local pc
                # pdf = pdfkit.from_string(html, filename_pdf, css=css)

                pdfs.append(filename_pdf)

            merger = PdfFileMerger()
            for pdf in pdfs:
                merger.append(pdf,import_bookmarks=False)
            merger.write("Raport_{}.pdf".format(kelas))
            merger.close()
            filenames = "Raport_{}.pdf".format(kelas)
            return redirect(url_for('download_rekap_nilai', filenames=filenames))
        elif cetak_semua =="0":
            if lihat == "0":
                # Cetak >> print the data of selected student (./nilai/Name_NISN_Raport.pdf)
                # print pdf
                filename_pdf = "{}_{}_Raport.pdf".format(nilai_sel[1], nisn_siswa)
                # filename_pdf = "Raport_{}_{}_{}.pdf".format(pelajaran,kelas,nilai_sel[1])
                headers_filename = "attachment; filename="+filename_pdf
                css = ["static/css/bootstrap.min.css","static/style.css"]
                ## uncomment config yang dipilih
                    # config for heroku :
                config = pdfkit.configuration(wkhtmltopdf='/app/bin/wkhtmltopdf')
                    # config for local ver 1 :
                # config = pdfkit.configuration(wkhtmltopdf='./bin/wkhtmltopdf')
                # config = pdfkit.configuration(wkhtmltopdf="/usr/local/bin/wkhtmltopdf")
                    # config for local Windows :
                # path_wkhtmltopdf = "C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe"
                # config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

                pdf = pdfkit.from_string(html, False,configuration=config, css=css)
                # config for local pc
                # pdf = pdfkit.from_string(html, False, css=css)


                response = make_response(pdf)
                response.headers["Content-Type"] = "application/pdf"
                response.headers["Content-Disposition"] = headers_filename
                response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
                response.headers["Pragma"] = "no-cache"
                response.headers["Expires"] = "0"
                response.headers['Cache-Control'] = 'public, max-age=0'
                return response
            elif lihat =="1":
                return html

        return render_template("rekap_wali.html",jlh_list=len(list_nisn), list_nisn=list_nisn, list_siswa=list_siswa,
                                    eval_type=eval_type, pelajaran=pelajaran, kelas=kelas)
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
            kelas = request.form['kelas']       #"VII"
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
        file_stream=stream_dropbox_file("/nilai/{}/{}/form_nilai_{}_{}.xlsx".format(folder_name_1, eval_type, pelajaran, kelas))
        nilai_siswa = pd.read_excel(file_stream)
        # show the list of siswa
        list_nisn = nilai_siswa["NISN"].tolist()
        list_siswa = nilai_siswa["Nama"].tolist()
        # load form komentar
        file_stream=stream_dropbox_file("/nilai/{}/{}/Komentar_{}_{}.xlsx".format(folder_name_1, eval_type, pelajaran, kelas))
        form_komentar = pd.read_excel(file_stream)
        form_komentar.set_index("NISN", inplace=True)

        if save_komentar=="1":
            # update Submit Komentar
            komentar = request.form['komentar']     # "Progres siswa sudah baik" # obtained from text box
            form_komentar.loc[(nisn_siswa), "Komentar"] = komentar
            # save komentar dataframe
            form_komentar.reset_index(inplace=True)
            form_komentar.to_excel("./nilai/Komentar_{}_{}.xlsx".format(pelajaran, kelas), index=None)
            with open("./nilai/Komentar_{}_{}.xlsx".format(pelajaran, kelas), 'rb') as f:
                dbx.files_upload(f.read(), "/nilai/{}/{}/Komentar_{}_{}.xlsx".format(folder_name_1, eval_type, pelajaran, kelas), mode=dropbox.files.WriteMode.overwrite)

        if nisn_siswa != None and save_komentar !="1":
            # present nilai siswa
            nilai_siswa_sel = nilai_siswa[nilai_siswa["NISN"] == nisn_siswa]
            aspek_materi = list(nilai_siswa_sel)
            # return str(nilai_siswa_sel)
            # return str(nisn_siswa)
            komentar_siswa_sel = form_komentar.loc[nisn_siswa,"Komentar"]
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
                                        pelajaran=pelajaran, kelas=kelas, semester=semester, aspek_materi=aspek_materi, komentar_siswa_sel=komentar_siswa_sel,
                                        eval_type=eval_type, nama_guru=session['nama_user'],jlh_aspek=len(aspek_materi))

        if cetak_semua == "1":
            # Cetak Semua >> iterating to print the data of all student (./nilai/Name_NISN_Mapel?.pdf)
            # merging pdf file
            pdfs = []
            for siswa in range(len(list_siswa)):
                nilai_siswa_sel = nilai_siswa[nilai_siswa["NISN"] == int(list_nisn[siswa])]
                aspek_materi = list(nilai_siswa_sel)
                # return str(nilai_siswa_sel)
                # return str(nisn_siswa)
                komentar_siswa_sel = form_komentar.loc[(list_nisn[siswa]),"Komentar"]
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
                                            pelajaran=pelajaran, kelas=kelas, semester=semester, aspek_materi=aspek_materi, komentar_siswa_sel=komentar_siswa_sel,
                                            eval_type=eval_type, nama_guru=session['nama_user'],jlh_aspek=len(aspek_materi))

                filename_pdf = "./nilai/{}_{}_{}.pdf".format(list_siswa[siswa], list_nisn[siswa], pelajaran)
                css = ["static/css/bootstrap.min.css","static/style.css"]
                ## uncomment config yang dipilih
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

                pdfs.append(filename_pdf)
            # for i in range(len(list_siswa)):
                # pdfs.append("./nilai/{}_{}_{}.pdf".format(list_siswa[i], list_nisn[i], pelajaran))
            merger = PdfFileMerger()
            for pdf in pdfs:
                merger.append(pdf,import_bookmarks=False)
            merger.write("Rekap_Nilai_{}_{}.pdf".format(pelajaran, kelas))
            merger.close()
            filenames = "Rekap_Nilai_{}_{}.pdf".format(pelajaran, kelas)
            return redirect(url_for('download_rekap_nilai', filenames=filenames))
        elif cetak_semua =="0":
            if lihat == "0":
                # Cetak >> print the data of selected student (./nilai/Name_NISN_Mapel?.pdf)
                # print pdf
                filename_pdf = "{}_{}_{}.pdf".format(nilai_sel[1], nilai_sel[0], pelajaran)
                # filename_pdf = "Raport_{}_{}_{}.pdf".format(pelajaran,kelas,nilai_sel[1])
                headers_filename = "attachment; filename="+filename_pdf
                css = ["static/css/bootstrap.min.css","static/style.css"]
                ## uncomment config yang dipilih
                    # config for heroku :
                config = pdfkit.configuration(wkhtmltopdf='/app/bin/wkhtmltopdf')
                    # config for local ver 1 :
                # config = pdfkit.configuration(wkhtmltopdf='./bin/wkhtmltopdf')
                # config = pdfkit.configuration(wkhtmltopdf="/usr/local/bin/wkhtmltopdf")
                    # config for local Windows :
                # path_wkhtmltopdf = "C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe"
                # config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

                pdf = pdfkit.from_string(html, False,configuration=config, css=css)
                # config for local pc
                # pdf = pdfkit.from_string(html, False, css=css)


                response = make_response(pdf)
                response.headers["Content-Type"] = "application/pdf"
                response.headers["Content-Disposition"] = headers_filename
                response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
                response.headers["Pragma"] = "no-cache"
                response.headers["Expires"] = "0"
                response.headers['Cache-Control'] = 'public, max-age=0'
                return response

            elif lihat == "1":
                return html

        # Reset Nilai >> go to step 4

        return render_template("rekap_mapel.html",jlh_list=len(list_siswa), list_nisn=list_nisn, list_siswa=list_siswa,
                                        eval_type=eval_type, pelajaran=pelajaran, kelas=kelas)

    else:
        return redirect(url_for("login"))

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


def unggah_form_nilai(file, eval_type, pelajaran, kelas):
    # grab upload excel name
    excel_name = file.filename
    file_destination = "/".join([path_nilai,excel_name])
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
    with open("./nilai/Komentar_{}_{}.xlsx".format(pelajaran, kelas), 'rb') as f:
        dbx.files_upload(f.read(), "/nilai/{}/{}/Komentar_{}_{}.xlsx".format(folder_name_1, eval_type, pelajaran, kelas), mode=dropbox.files.WriteMode.overwrite)

    form_nilai = form_nilai.iloc[:,:-1]
    col_title = form_nilai.columns.tolist()
    list_pengetahuan = []; list_keterampilan = []
    for title in col_title:
        tmp = title.split("_")
        if len(tmp) > 1:
            if tmp[1] == "Pengetahuan":
                list_pengetahuan.append(title)
            elif tmp[1] == "Keterampilan":
                list_keterampilan.append(title)
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
    with open(file_destination, 'rb') as f:
        dbx.files_upload(f.read(), "/nilai/{}/{}/form_nilai_{}_{}.xlsx".format(folder_name_1, eval_type, pelajaran, kelas), mode=dropbox.files.WriteMode.overwrite)
    # update input status
    # load status file
    file_stream=stream_dropbox_file("/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type))
    status_nilai = pd.read_excel(file_stream)
    status_nilai.set_index("Mata Pelajaran", inplace=True)
    status_nilai.loc[pelajaran, kelas] = 1
    status_nilai.to_excel("./tmp/status_nilai.xlsx")
    with open("./tmp/status_nilai.xlsx", 'rb') as f:
        dbx.files_upload(f.read(), "/nilai/{}/{}/status_nilai.xlsx".format(folder_name_1, eval_type), mode=dropbox.files.WriteMode.overwrite)

    # copy list guru mapel
    try:
        dbx.files_get_metadata("/nilai/{}/{}/guru_mapel.xlsx".format(folder_name_1, eval_type))
    except:
        dbx.files_copy("/guru_mapel.xlsx", "/nilai/{}/{}/guru_mapel.xlsx".format(folder_name_1, eval_type))

def save_template_form_nilai(pelajaran, kelas, aspek_materi, eval_type):
    # 5. Unduh dan unggah form nilai
    # load data siswa
    file_stream=stream_dropbox_file("/data_siswa.xlsx")
    data_siswa = pd.read_excel(file_stream)
    # define aspek penilaian
    aspek_lain = ["Spiritual_Predikat", "Sosial_Predikat"]
    aspek_materi = aspek_materi.split(";")
    tmp = aspek_materi[:]
    aspek_materi = []
    for materi in tmp:
        materi = materi.split(" ")
        if "" in materi:
            materi.remove("")
        materi = ' '.join(materi)
        aspek_materi.append(materi)
    # aspek_materi.extend([eval_type, "Nilai Akhir"])
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
    # form_nilai.drop(columns=["{}_Keterampilan".format(eval_type)], inplace=True)
    # for aspek in ["Catatan Guru Mapel"]:
    #     form_nilai[aspek] = ""
    form_nilai.to_excel("tmp/form_nilai_{}_{}.xlsx".format(pelajaran, kelas), index=None)
    # send the file to user
    tahun_ajaran, semester, folder_name_1 = check_period()
    # create file komentar
    # create komentar dataframe
    # komentar = 0
    # try:
    #     dbx.files_get_metadata("/nilai/{}/{}/Komentar_{}_{}.xlsx".format(folder_name_1, eval_type, pelajaran, kelas))
    # except:
    #     komentar = 1
    # if komentar:
    #     form_komentar = data_siswa[data_siswa["Kelas"] == kelas][["NISN", "Nama"]]
    #     form_komentar["Komentar"] = ""
    #     form_komentar.to_excel("./nilai/Komentar_{}_{}.xlsx".format(pelajaran, kelas), index=None)
    #     with open("./nilai/Komentar_{}_{}.xlsx".format(pelajaran, kelas), 'rb') as f:
    #         dbx.files_upload(f.read(), "/nilai/{}/{}/Komentar_{}_{}.xlsx".format(folder_name_1, eval_type, pelajaran, kelas), mode=dropbox.files.WriteMode.overwrite)

@app.route("/clear")
def clearSession():
    session.pop('user',None)
    session.pop('nama_user',None)
    return redirect(url_for('login'))

@app.route('/input/<string:filenames>')
def download_rekap_nilai(filenames):
    # send the merged pdf file to user
    return send_from_directory(APP_ROOT, filename=filenames, as_attachment=True)

@app.route('/input/<string:pelajaran>/<string:kelas>')
def download_template_nilai(pelajaran, kelas):
    path_data = os.path.join(APP_ROOT, 'tmp/')
    filenames = "form_nilai_{}_{}.xlsx".format(pelajaran, kelas)
    # send the file to user
    return send_from_directory(path_data, filename=filenames, as_attachment=True)

@app.route('/admin/data-siswa/template')
def download_template_data_siswa():
    path_data = os.path.join(APP_ROOT, 'data/')
    filenames = "data_siswa_template.xlsx"
    # return filenames
    # send the file to user
    return send_from_directory(path_data, filename=filenames, as_attachment=True)

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
