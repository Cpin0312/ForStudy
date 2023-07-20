UPDATE m_kaiin_kihon SET kaiin_pswd = UPPER(kaiin_pswd), upd_nitiji = '20191211173000000',  upd_usr = 'SYSTEM';



Select
*
FRom
m_kaiin_kihon
where
upd_usr = 'SYSTEM'
