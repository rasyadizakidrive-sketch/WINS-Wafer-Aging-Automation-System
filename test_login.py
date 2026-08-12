from SAP.sap_launcher import launch_sap

from SAP.sap_login import login_sap

from config import *

launch_sap()

login_sap(

    SAP_CLIENT,

    SAP_USERNAME,

    SAP_PASSWORD,

    SAP_LANGUAGE

)