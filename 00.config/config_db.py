import platform

if platform.system() == 'Windows':
    DB_PATH = r"C:\Users\LISARR\OneDrive - Salvesen Logística S.A\00.DB\2026.db"
elif platform.system() == 'Darwin':
    DB_PATH = "/Volumes/RR/DB/inform_27.db"