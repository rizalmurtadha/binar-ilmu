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

def check_predikat(avg):
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
    return predikat

# 1. Halaman login
# login to dropbox
token = "UUqaOMObW8sAAAAAAAAAAcvkmOaxYzJs7ZRhjCRDqVqvKuP-8Gd1W0n7i6CjhKNK"
dbx = dropbox.Dropbox(token)

@app.route("/",methods=["GET", "POST"])
def login():
    if "user" in session:
        return redirect(url_for("role"))
    else:
        if request.method=="POST":
            try:
                Login = request.form['Login']
            except:
                Login = "0"
            if(Login=="1"):
                # load data guru
                file_stream=stream_dropbox_file("/data_guru.xlsx")
                data_guru = pd.read_excel(file_stream)
                # dbx.files_download_to_file("./tmp/data_guru.xlsx", "/data_guru.xlsx")
                # data_guru = pd.read_excel("./tmp/data_guru.xlsx")
                # verifikasi akun
                # misal input id_guru di-assign sebagai variable "id_guru"
                id_guru = int(request.form['id_guru'])
                # misal input password di-assign sebagai variable "input_passwd"
                input_passwd = request.form['password']
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
                        for aspek in ["Sikap", "Keterampilan", "Pengetahuan"]:
                            form_nilai["{}_{}".format(mpl, aspek)] = form_nilai_mapel[aspek].values
                form_nilai["Rata_rata"] = form_nilai.iloc[:,1:].mean(axis=1).values
                form_nilai["Predikat"] = form_nilai["Rata_rata"].apply(lambda x: check_predikat(x))
                # file_rekap_nilai = "Rekap_Nilai_{}.xlsx".format(kelas)
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
                    # file_komentar ="Komentar_{}.xlsx".format(kelas) 
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
                save_template_form_nilai(pelajaran, kelas, aspek_materi)
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
        lihat = None
        if request.method=="POST":
            eval_type = request.form['eval_type']   #"PTS" 
            pelajaran = request.form['pelajaran']   #"Matematika"
            kelas = request.form['kelas']       #"VII"
            try:
                save_komentar = request.form['save_komentar']
            except:
                save_komentar="0"
            # Tombol Lihat
            try:
                lihat = request.form['lihat']
                nisn_siswa = int(request.form['nisn_siswa']) # 72100585 obtained from previous step
            except:
                lihat="0"
                nisn_siswa = None
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
            for i in range(nilai_siswa_sel.shape[1]):
                if i > 2 and i < nilai_siswa_sel.shape[1]-3 and i%3==0 :
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
            
            jlh_mapel = int((len(nilai_sel)-4)/3)
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
                    if i > 2 and i < nilai_siswa_sel.shape[1]-3 and i%3==0 :
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
                
                jlh_mapel = int((len(nilai_sel)-4)/3)
                html= render_template("nilai_wali.html",nilai_sel=nilai_sel, tahun_ajaran=tahun_ajaran, cetak="0",lihat=lihat, nisn_siswa=list_nisn[siswa],
                                            pelajaran=pelajaran, kelas=kelas, semester=semester, nama_mapel=nama_mapel, komentar_siswa_sel=komentar_siswa_sel,
                                            eval_type=eval_type, nama_guru=session['nama_user'],jlh_mapel=jlh_mapel,jlh_nilai_sel=len(nilai_sel))

                filename_pdf = "./nilai/{}_{}_Raport.pdf".format(list_siswa[siswa], list_nisn[siswa])
                css = ["static/css/bootstrap.min.css","static/style.css"]
                ## uncomment config yang dipilih
                    # config for heroku :
                config = pdfkit.configuration(wkhtmltopdf='/usr/local/bin/wkhtmltopdf')
                    # config for local ver 1 :
                # config = pdfkit.configuration(wkhtmltopdf='./bin/wkhtmltopdf')
                    # config for local ver 2 :
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
                config = pdfkit.configuration(wkhtmltopdf='/usr/local/bin/wkhtmltopdf')
                    # config for local ver 1 :
                # config = pdfkit.configuration(wkhtmltopdf='./bin/wkhtmltopdf')
                    # config for local ver 2 :
                path_wkhtmltopdf = "C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe"
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
            # aspek_materi=request.args.get('aspek_materi')
        cetak_semua = "0"
        lihat = None
        if request.method=="POST":
            eval_type = request.form['eval_type']   #"PTS" 
            pelajaran = request.form['pelajaran']   #"Matematika"
            kelas = request.form['kelas']       #"VII"
            # Tombol Lihat
            try:
                lihat = request.form['lihat']
                nisn_siswa = int(request.form['nisn_siswa']) # ex:72100585  obtained from previous step
            except:
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

        if nisn_siswa != None and cetak_semua !="1":
            # present nilai siswa
            nilai_siswa_sel = nilai_siswa[nilai_siswa["NISN"] == nisn_siswa]
            aspek_materi = list(nilai_siswa_sel)
            batas = len(aspek_materi) - 4
            aspek_materi = [aspek_materi[i] for i in range(6,batas) ]
            nilai_sel = []
            
            for i in range(nilai_siswa_sel.shape[1]):
                if i > 5 and i < nilai_siswa_sel.shape[1]-4 or i == nilai_siswa_sel.shape[1]-1:
                    nilai_sel.append("{0:.2f}".format(nilai_siswa_sel.iloc[0,i]))
                else:
                    nilai_sel.append(nilai_siswa_sel.iloc[0,i])

            html= render_template("nilai_mapel.html",nilai_sel=nilai_sel, tahun_ajaran=tahun_ajaran, cetak="0",lihat=lihat,
                                        pelajaran=pelajaran, kelas=kelas, semester=semester, aspek_materi=aspek_materi,
                                        eval_type=eval_type, nama_guru=session['nama_user'],jlh_aspek=len(aspek_materi))

        if cetak_semua == "1":
            # Cetak Semua >> iterating to print the data of all student (./nilai/Name_NISN_Mapel?.pdf)
            # merging pdf file
            pdfs = []
            for siswa in range(len(list_siswa)):
                nilai_siswa_sel = nilai_siswa[nilai_siswa["NISN"] == int(list_nisn[siswa])]
                aspek_materi = list(nilai_siswa_sel)
                batas = len(aspek_materi) - 4
                aspek_materi = [aspek_materi[i] for i in range(6,batas) ]
                nilai_sel = []
                for i in range(nilai_siswa_sel.shape[1]):
                    if i > 5 and i < nilai_siswa_sel.shape[1]-4 or i == nilai_siswa_sel.shape[1]-1:
                        nilai_sel.append("{0:.2f}".format(nilai_siswa_sel.iloc[0,i]))
                    else:
                        nilai_sel.append(nilai_siswa_sel.iloc[0,i])
                html= render_template("nilai_mapel.html",nilai_sel=nilai_sel, tahun_ajaran=tahun_ajaran, cetak="0",lihat=lihat,
                                            pelajaran=pelajaran, kelas=kelas, semester=semester, aspek_materi=aspek_materi,
                                            eval_type=eval_type, nama_guru=session['nama_user'],jlh_aspek=len(aspek_materi))

                filename_pdf = "./nilai/{}_{}_{}.pdf".format(list_siswa[siswa], list_nisn[siswa], pelajaran)
                css = ["static/css/bootstrap.min.css","static/style.css"]
                ## uncomment config yang dipilih
                    # config for heroku :
                config = pdfkit.configuration(wkhtmltopdf='/usr/local/bin/wkhtmltopdf')
                    # config for local ver 1 :
                # config = pdfkit.configuration(wkhtmltopdf='./bin/wkhtmltopdf')
                    # config for local ver 2 :
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
                config = pdfkit.configuration(wkhtmltopdf='/usr/local/bin/wkhtmltopdf')
                    # config for local ver 1 :
                # config = pdfkit.configuration(wkhtmltopdf='./bin/wkhtmltopdf')
                    # config for local ver 2 :
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
    form_nilai["Pengetahuan"] = 0
    for i in range(form_nilai.shape[0]):
        tmp = np.average(form_nilai.iloc[:,6:-4].values)
        form_nilai.loc[i,"Pengetahuan"] = tmp
    form_nilai.to_excel(file_destination, index=None)
    # load form nilai
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

def save_template_form_nilai(pelajaran, kelas, aspek_materi):
    # 5. Unduh dan unggah form nilai
    # load data siswa
    file_stream=stream_dropbox_file("/data_siswa.xlsx")
    data_siswa = pd.read_excel(file_stream)
    # define aspek penilaian
    aspek_lain = ["Spiritual_Predikat", "Spiritual_Deskripsi", "Sosial_Predikat", "Sosial_Deskripsi"]
    aspek_materi = aspek_materi.split(";")
    # Unduh form nilai
    # create dataframe of form penilaian
    form_nilai = data_siswa[data_siswa["Kelas"] == kelas][["NISN", "Nama"]]
    for aspek in aspek_lain:
        form_nilai[aspek] = ""
    for aspek in aspek_materi:
        form_nilai[aspek] = ""
    for aspek in ["Sikap", "Keterampilan", "Komentar"]:
        form_nilai[aspek] = ""
    form_nilai.to_excel("tmp/form_nilai_{}_{}.xlsx".format(pelajaran, kelas), index=None)


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
