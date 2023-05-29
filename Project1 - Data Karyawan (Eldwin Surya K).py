from datetime import datetime
import tabulate
import pandas as pd

menu = 0
menu_view = 0

karyawan = [{"NPK":14878, "Nama_Karyawan":"Eldwin Surya", "Email":"eldwinsk@gmail.com", "No_HP":"081228039056","Tgl_Join":"23/03/2023","Status":"Aktif"},
            {"NPK":13981, "Nama_Karyawan":"Johan Setya", "Email":"johan@gmail.com", "No_HP":"081889332012","Tgl_Join":"26/05/2023","Status":"Aktif"},
            {"NPK":12331, "Nama_Karyawan":"Ary Bambang", "Email":"ary@gmail.com", "No_HP":"081775883229","Tgl_Join":"03/11/2020","Status":"Aktif"},
            {"NPK":14331, "Nama_Karyawan":"Rahmadi Ram", "Email":"rahmadi@gmail.com", "No_HP":"081849998221","Tgl_Join":"03/11/2022","Status":"Non Aktif"},
            {"NPK":11441, "Nama_Karyawan":"Sophie Latih", "Email":"sophie@gmail.com", "No_HP":"081112339857","Tgl_Join":"11/08/2018","Status":"Aktif"},
            {"NPK":11721, "Nama_Karyawan":"Ivan Oktona", "Email":"ivan@gmail.com", "No_HP":"081447821390","Tgl_Join":"01/10/2019","Status":"Aktif"},
            {"NPK":15392, "Nama_Karyawan":"Eka Putri", "Email":"eka@gmail.com", "No_HP":"081338229432","Tgl_Join":"05/01/2023","Status":"Non Aktif"}]
headers = karyawan[0].keys()

#Function untuk menampilkan Data Karyawan
def menampilkan_daftar_karyawan():
    print("\nDaftar Karyawan\n")
    rows = [x.values() for x in karyawan]
    print (tabulate.tabulate(rows, headers))


def menampilkan_daftar_karyawan_aktif():
    print("\nDaftar Karyawan Aktif\n")
    rows = [x.values() for x in karyawan if x['Status']=="Aktif"]
    print (tabulate.tabulate(rows, headers))

def menampilkan_daftar_karyawan_nonaktif():
    print("\nDaftar Karyawan Non Aktif\n")
    rows = [x.values() for x in karyawan if x['Status']=="Non Aktif"]
    print (tabulate.tabulate(rows, headers))

    
#While di Menu Utama, hanya keluar apabila user input Exit
while menu!=6:
    print("\n\tMain Menu Karyawan")
    print("1. View Data Karyawan")
    print("2. Menambahkan Data Karyawan")
    print("3. Mengubah Data Karyawan")
    print("4. Menghapus Data Karyawan")
    print("5. Export Data Karyawan to Excel")
    print("6. Exit")
    #Pengecekan user hanya bisa memasukan angka yang ada di menu utama
    try:
        menu = int(input("Masukan angka menu yang akan dijalankan = "))
    except ValueError:
        print("Hanya inputkan angka yang ada di Menu")
  
    #If di Menu View Data Karyawan
    if(menu==1):
        print("1. View Karyawan Aktif")
        print("2. View Karyawan Non Aktif")
        print("3. View Seluruh Data Karyawan")
        print("4. Kembali ke menu sebelumnya")
        try:
            menu_view = int(input("Masukan angka Menu yang ingin dijalankan = "))
        except ValueError:
            print("Hanya inputkan angka yang ada di Menu")
        if(menu_view == 1):
            menampilkan_daftar_karyawan_aktif()
        if(menu_view == 2):
            menampilkan_daftar_karyawan_nonaktif()
        if(menu_view == 3):
            menampilkan_daftar_karyawan()
        if(menu_view == 4):
            pass
    
    #If di Menu Menambah Data Karyawan
    if(menu==2):
        i_npk = 0
        i_email = 0
        index_email = 0
        i_nohp = 0
        i_join = 0
        menampilkan_daftar_karyawan()
        print("\nMenambahkan Data Karyawan")
        #Pengecekan NPK yang dimasukan sudah benar 5 digit dan belum ada di database
        while i_npk==0:
            try:
                NPK=int(input("Masukan 5 Digit Nomor Induk Karyawan = "))
                for i in range(len(karyawan)):
                    if(karyawan[i]['NPK']==NPK):
                        print("NPK yang anda masukan sudah ada di database")
                        break
                    else:
                        if(len(str(NPK))!=5):
                            print("NPK yang anda masukan kurang/lebih dari 5 digit")
                            break
                        else:
                            i_npk=1
            #Pengecekan Error Inputan selain digit angka
            except ValueError:
                print("Hanya Masukan 5 digit angka")
            
        Nama=input("Masukan Nama Karyawan = ")
        #Pengecekan format by email
        while i_email==0:
            Email=input("Masukan Email Karyawan (Menggunakan @gmail.com)= ")
            index_email = Email.find("@gmail.com")
            if(index_email>0):
                i_email=1
            else:
                print("Format Email yang anda masukan salah, mohon menggunakan @gmail.com")
        
        #Pengecekan inputan Nomor HP, apakah diantara 10-14 digit angka(No HP Indonesia)
        while i_nohp==0:
            No_Hp=input("Masukan Nomor HP Karyawan = ")
            if(len(No_Hp)<10 or len(No_Hp)>14):
                print("Masukan Nomor Handphone antara 10 hingga 14 digit")
            else:
                if(not No_Hp.isdigit()):
                    print("Hanya masukan digit angka")
                else:
                    i_nohp=1

        #Pengecekan inputan format tanggal join dd/mm/yyyy
        while i_join ==0:
            Tgl_Join=input("Masukan Tanggal Join Perusahaan (dd/mm/yyyy)= ")

            isValidDate = True
            try:
                datetime.strptime(Tgl_Join,"%d/%m/%Y")
            except ValueError:
                isValidDate = False

            if(isValidDate):
                i_join=1
            else:
                print("Format Tanggal yang anda masukan salah (dd/mm/yyyy)")
    
        karyawan_baru = {"NPK":NPK, "Nama_Karyawan":Nama, "Email":Email, "No_HP":No_Hp,"Tgl_Join":Tgl_Join,"Status":"Aktif"}
        karyawan.append(karyawan_baru)
        menampilkan_daftar_karyawan_aktif()

    #If Menu Utama Perubahan/Update Data Karyawan
    if(menu==3):
        i_npk = 0
        i_email = 0
        index_email = 0
        i_nohp = 0
        i_join = 0
        menampilkan_daftar_karyawan()
        print("\nMengubah Data Karyawan")
        while i_npk==0:
            try:
                ubah_NPK = int(input("Masukan NPK Karyawan yang ingin diubah = "))
                #Pengecekan inputan NPK yang sudah sama dengan NPK database
                for i in range (len(karyawan)):
                    if(ubah_NPK==karyawan[i]["NPK"]):
                        i_npk=1
                if(i_npk==0):
                    print("NPK yang anda masukan salah")
            except ValueError:
                print("Hanya Masukan 5 digit angka")
    
        print("Pilih opsi perubahan : ")
        print("[1] Ubah Nama Karyawan")
        print("[2] Ubah Email Karyawan")
        print("[3] Ubah No HP Karyawan")
        print("[4] Ubah Status Karyawan")
        try:
            ubah_opsi = int(input("Masukan opsi perubahan = "))
        except ValueError:
            print("Hanya inputkan angka yang ada di Menu")

        if(ubah_opsi==1):
            for i in range (len(karyawan)):
                if(ubah_NPK == karyawan[i]['NPK']):
                    print("NPK Karyawan yang ingin diubah : {}".format(ubah_NPK))
                    print("Nama Karyawan : {}".format(karyawan[i]['Nama_Karyawan']))
                    nama_baru = input("Masukan Nama Baru Karyawan : ")
                    karyawan[i]['Nama_Karyawan'] = nama_baru

        if(ubah_opsi==2):
            for i in range (len(karyawan)):
                if(ubah_NPK == karyawan[i]['NPK']):
                    while i_email==0:
                        print("NPK Karyawan yang ingin diubah : {}".format(ubah_NPK))
                        print("Email Karyawan : {}".format(karyawan[i]['Email']))
                        Email_Baru = input("Masukan Email Baru Karyawan : ")
                        index_email = Email_Baru.find("@gmail.com")
                        if(index_email>0):
                            i_email=1
                        else:
                            print("Format Email yang anda masukan salah, mohon menggunakan @gmail.com")
                    karyawan[i]['Email'] = Email_Baru

        if(ubah_opsi==3):
            for i in range (len(karyawan)):
                if(ubah_NPK == karyawan[i]['NPK']):
                    print("NPK Karyawan yang ingin diubah : {}".format(ubah_NPK))
                    print("No_HP Karyawan : {}".format(karyawan[i]['No_HP']))
                    while i_nohp==0:
                        No_Baru=input("Masukan Nomor Hp Baru Karyawan = ")
                        if(len(No_Baru)<10 or len(No_Baru)>14):
                            print("Masukan Nomor Handphone antara 10 hingga 14 digit")
                        else:
                            if(not No_Baru.isdigit()):
                                print("Hanya masukan digit angka")
                            else:
                                i_nohp=1
                    karyawan[i]['No_HP'] = No_Baru

        if(ubah_opsi==4):
            for i in range(len(karyawan)):
                if(ubah_NPK == karyawan[i]['NPK']):
                    print("NPK Karyawan yang ingin diubah : {}".format(ubah_NPK))
                    if( karyawan[i]['Status']=="Aktif"):
                        Status_Baru = "Non Aktif"
                    else:
                        Status_Baru = "Aktif"
                    print("Status awal : {} \nTelah diubah menjadi Status : {}".format(karyawan[i]['Status'],Status_Baru))
                    karyawan[i]['Status']=Status_Baru
    
    #If Menghapus Data Karyawan 
    if(menu==4):
        i_npk=0
        print("Hanya Bisa Menghapus Data Karyawan Non Aktif")
        menampilkan_daftar_karyawan_nonaktif()
        while i_npk==0:
            try:
                npk_hapus = int(input("Masukan NPK Data Karyawan yang ingin di hapus : "))
                for i in range(len(karyawan)):
                    #Pengecekan hanya NPK yang statusnya Non Aktif, yang hanya bisa di hapus oleh user
                    if(npk_hapus == karyawan[i]['NPK'] and karyawan[i]['Status']=="Non Aktif"):
                        print("Data NPK {} telah berhasil di hapus".format(npk_hapus))
                        del karyawan[i]
                        i_npk=1
                        break
                    else:
                        pass
                if(i_npk==0):
                    print("NPK yang anda masukan salah, mohon menggunakan Data NPK Non Aktif")
            except ValueError:
                print("Hanya masukan 5 digit angka")

    #If convert data all Karyawan to Excel
    if(menu==5):
        cetak = pd.DataFrame(karyawan)
        cetak_to_excel = pd.ExcelWriter("Data Karyawan.xlsx")
        cetak.to_excel(cetak_to_excel, sheet_name="List Data")

        cetak_to_excel.close()
        print("File Data Karyawan berhasil di Export ke dalam Excel")



print("\nProgram Data Karyawan Log out, Terima Kasih\n")
