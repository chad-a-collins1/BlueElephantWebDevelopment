<%
' App Constants
'*********************************************
Const GLOB_DEBUG = "yes"
Const GLOB_PRODUCTION = "no"
'Const ALLOW_ASPQMAIL = "no"
'Const ALLOW_SSL = "no"
Const HREF_SSL = "https://bellerophon.ntsecure.net/blueelephantwebdev/"
Const HREF_TEST_HOME = "http://localhost/blueelep/"
Const HREF_PROD_HOME = "http://www.blueelephantwebdev.com"

Const PAGE_HOME = "index.htm"
Const PAGE_RATES = "Rates.htm"
Const PAGE_ABOUT = "About.htm"
Const PAGE_SERVICES = "Services.htm"
Const PAGE_LOGIN = "Login.asp"
Const PAGE_INITIALCONTACT = "InitialContact.asp"
Const PAGE_PORTAL = "Portal.asp"
Const PAGE_ERROR = "Error.asp"
Const PAGE_500100ERROR = "500-100.asp"
Const FOLDER_UPLOAD = "uploads"

Const PIC_MAIN = "main.jpg"
Const PIC_CONTACT = "contact.jpg"
Const PIC_SERVICES = "services.jpg"
Const PIC_LOGIN = "login.jpg"
Const PIC_RATES = "rates.jpg"
Const PIC_ABOUT = "about.jpg"
Const PIC_PORTAL = "portal.jpg"
Const PIC_ERROR = "services.jpg"

Const TITLE_MAIN = "BlueElephantWebDev.com Home"
Const TITLE_CONTACT = "BlueElephantWebDev.com Contact Us"
Const TITLE_SERVICES = "BlueElephantWebDev.com Services"
Const TITLE_LOGIN = "BlueElephantWebDev.com Client Login"
Const TITLE_RATES = "BlueElephantWebDev.com Our Rates"
Const TITLE_ABOUT = "BlueElephantWebDev.com About Us"
Const TITLE_PORTAL = "BlueElephantWebDev.com Client Portal"
Const TITLE_ERROR = "BlueElephantWebDev.com Error"

Const LOGO_GRAPHIC = "graphic_logo-bar.js"
Const LOGO_FLASH = "logo-bar.js"

Const COLOR_BTN1 = "#94BED6"
Const COLOR_BTN1_DN = "#267EAE"

Const ACTION_LOGIN_VALIDATE = "alogv_574hff"
Const ACTION_VALIDATE = "av_2984wfy"
Const ACTION_SIGNUP_VALIDATE = "asigv_483fbf"
Const ACTION_LOGIN = "alog_7fg4"
Const ACTION_MAIN = "am_576fhhd"
Const ACTION_SUCCESS = "ascs_56fyd"
Const ACTION_MAIN_VALIDATE_PAGESELECT = "amvps_576fhhd"
Const ACTION_MAINSUB = "ams_qgn574"
Const ACTION_MAINSUB_AFTER = "amsa_5e574"
Const ACTION_ADDNEW = "aan_5784hdh"
Const ACTION_VIEW = "av_hf823"
Const ACTION_VIEW_FROM_SUB = "avfs_3ughef"
Const ACTION_VIEW_CANCEL = "avc_h6y3"
Const ACTION_EDIT = "ae_hg743"
Const ACTION_DELETE = "ad_gfh4637"
Const ACTION_VALIDATE_ADDNEW = "avan_ghty56473"
Const ACTION_VALIDATE_EDIT = "ave_459u5ht"
Const ACTION_VALIDATE_DELETE = "avd_45h3y"
Const QSVAR = "x"
Const ACTION_LOGOUT = "alo_t7geb"
Const ACTION_LOGOUT_GOHOME = "algh_t895g"
Const ACTION_LOGOUT_NORMAL = "aln_gfjhd97"
Const ACTION_LOGOUT_GOLOGIN = "algl_3t8hf"


Const USERNAME_MIN_LENGTH = 5
Const USERNAME_MAX_LENGTH = 50
Const PASSWORD_MIN_LENGTH = 8
Const PASSWORD_MAX_LENGTH = 15

Const ERR_INVALID_PWD = 998
Const ERR_INVALID_UID = 999 
Const ERR_INVALID_SESSION_ID = 902
Const ERR_INVALID_LOGIN_DATE = 903
Const ERR_INVALID_ACTION = 916
Const ERR_INVALID_LAST_PAGE = 917
Const ERR_NO_LOGIN = 918
Const ERR_UNKNOWN = 919
Const ERR_INVALID_ERR = 927


Const ERR_GET_USERDATA_SPEC = 904
Const ERR_GET_USERDATA_DESC = 905
Const ERR_ADDNEW_USERDATA = 906
Const ERR_EDIT_USERDATA = 907
Const ERR_DELETE_USERDATA = 924
Const ERR_DBINSERT_USERLOGIN = 920
Const ERR_DBDELETE_USERDATA = 921
Const ERR_INVALID_VAP_RC = 922
Const ERR_DBINSERT_USERLOGOUT = 923
Const ERR_NO_USERDATA_SPEC_INCOOK = 925
Const ERR_NO_USERDATASPEC = 925
Const ERR_UNKNOWN_VALERR = 926
Const ERR_INVALID_SUBMIT = 928
Const ERR_GET_CALENDARMONTH_FROM_DB = 929
Const ERR_INVALID_CURDATE = 930
Const ERR_GET_MONTHINFO_FROM_DB = 931
Const ERR_GET_MONTHINFO_FROM_COOKIE = 932
Const ERR_INVALID_CALENDARFREQ = 933
Const ERR_GET_TIMEZONE = 934
Const ERR_INVALID_USERAGENT = 935
Const ERR_INVALID_CASE = 936
Const ERR_GET_STATES = 937
Const ERR_GET_CCTYPE = 938
Const ERR_INVALID_SERVICETYPE = 939
Const ERR_GET_DCCOLOR = 940
Const ERR_ACCESS_VIOLATION = 941
Const ERR_GET_ACCTTYPES = 942
Const ERR_SENDEMAIL_ACCTINFO = 943
Const ERR_GET_ACCTINFO = 944
Const ERR_SEND_EMAIL = 945
Const ERR_SEND_FTEMAIL = 946
Const ERR_NO_RM_ID = 947
Const ERR_INVALID_DELIV = 948
Const ERR_NO_PROSPECT = 949
Const ERR_UNAUTH_DOWNLOAD = 950

Const USERAGENT_IE = "ua_msie"
Const USERAGENT_NN = "ua_netscape"
Const USERAGENT_OTHER = "ua_other"

Const LEN_DATETIME = 19     ' YYYY-MM-DD HH:MM:SS
Const LOGIN_TIMEOUT = 1200   ' 1200 seconds



Const VALERR_NO_USERNAME = 700
Const VALERR_NO_PASSWORD = 701
Const VALERR_NO_EMAIL = 702
Const VALERR_NO_QUESTION = 703
Const VALERR_NO_ANSWER = 704
Const VALERR_INVALID_USERNAME = 705
Const VALERR_INVALID_PASSWORD = 706
Const VALERR_INVALID_EMAIL = 707
Const VALERR_CONFIRM_PASSWORD = 708
Const VALERR_CONFIRM_EMAIL = 709
Const VALERR_NO_AGREE = 710
Const VALERR_NO_PACKAGE = 711
Const VALERR_NO_BFIRSTNAME = 712
Const VALERR_NO_BLASTNAME = 713
Const VALERR_NO_BSTREET = 714
Const VALERR_NO_BCITY = 715
Const VALERR_NO_BSTATE = 716
Const VALERR_NO_BZIP = 717
Const VALERR_NO_CCTYPE = 718
Const VALERR_NO_CCNUMBER = 719
Const VALERR_NO_CCNAME = 720
Const VALERR_NO_CCEXP = 721
Const VALERR_INVALID_BZIP = 722  
Const VALERR_NO_TONAME = 723 
Const VALERR_NO_FROMNAME = 724 
Const VALERR_NO_MESSAGE1 = 725 
Const VALERR_NO_PEMAIL = 726
Const VALERR_NO_DELIVERY = 727
Const VALERR_INVALID_PEMAIL = 728 
Const VALERR_CONFIRM_PEMAIL = 729
Const VALERR_NO_REMAIL = 730
Const VALERR_INVALID_REMAIL = 731
Const VALERR_CONFIRM_REMAIL = 732
Const VALERR_INVALID_ACODE = 733

Const VALERROR_DUPE_USERNAME = 223
Const VALERROR_DUPE_EMAIL = 224




%>