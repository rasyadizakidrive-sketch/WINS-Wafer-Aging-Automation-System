import config
from datetime import datetime, timedelta

from SAP.mb52_runner import run_mb52
from SAP.sap_reader import get_total_stock


DIVISOR = {
    "M6": 850000,
    "G12": 110000,
}


def calculate_doi(product):

    config.set_product(product)

    print(f"\n[INFO] Calculating DOI ({product})...")

    run_mb52()

    total_qty = get_total_stock()

    doi = round(
        total_qty / DIVISOR[product],
        1
    )

    #print(f"[OK] {product} Total Qty : {total_qty:,}")
    print(f"[OK] DOI : {doi} Days")
    
    return total_qty, doi


def run_doi():

    _, doi_m6 = calculate_doi("M6")

    _, doi_g12 = calculate_doi("G12")

    last_updated=datetime.now()

    return {
        'last_updated':last_updated,
        'm6':{'doi':doi_m6,'until':last_updated+timedelta(days=doi_m6)},
        'g12':{'doi':doi_g12,'until':last_updated+timedelta(days=doi_g12)},
    }

if __name__ == "__main__":

    result = run_doi()

    print()
    print("==============================")
    print(f"M6  DOI : {result['m6']['doi']} Days")
    print(f"G12 DOI : {result['g12']['doi']} Days")
    print("==============================")