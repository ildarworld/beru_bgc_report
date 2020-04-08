import pytest
from decimal import Decimal
from app.app import BeruAnalizer, Payment

def test_loadFile():
    ba = BeruAnalizer('app/tests/data/orders.xlsx')
    assert len(ba._orders) > 10
    assert len(ba._bgc) > 10


def test_download():
    ba = BeruAnalizer('app/tests/data/orders.xlsx')
    ba.getBGCExcel()
    assert True == True


def test_bgc_payment():
    ba = BeruAnalizer('app/tests/data/orders.xlsx')
    bgc = ba._bgc[295712]
    assert 29196.0 == bgc.paymentSum


def test_bgc_payment_item():
    ba = BeruAnalizer('app/tests/data/orders.xlsx')
    bgc = ba._bgc[295712]

    item = [item.cost for item in bgc.items if item.sku == '00008752']
    print(item)
    assert bgc.items[2].cost < 0 


def test_bgc_payment_creation():
    claim_payment = Payment(
        credit=179.0,
        BGC=2121821,
        BGC_date='date',
        payment_id='',
        payment_ladger_date=''
    )
    print(claim_payment)
    assert claim_payment.payment < 0
