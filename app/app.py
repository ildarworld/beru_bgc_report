import pandas as pd
import numpy as np
from decimal import Decimal
from datetime import datetime


def strToDate(str):
    return datetime.strptime(str, '%d.%m.%Y')


def get_col_widths(dataframe):
    # First we find the maximum length of the index column   
    idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]



class Payment:

    def __init__(self, BGC, BGC_date, payment_id, payment_ladger_date, debit=0, credit=0):
        self._debit = debit
        self._credit = credit
        self.BGC = BGC
        self.BGC_date = BGC_date
        self.payment_id = payment_id
        self.payment_ladger_date = payment_ladger_date

    @property
    def payment(self):
        return self._debit - self._credit


# Платежное поручение
class BGC:

    def __init__(self, number, date, item):
        self.number = number
        self.date = date
        self.items = [item]

    @property
    def paymentSum(self):
        return sum([s.cost for s in self.items])

    def addItem(self, item):
        self.items.append(item)


# Позиция платежного поручения
class BGCItem:
    
    def __init__(self, sku, qty, cost, order):
        self.sku = sku
        self.qty = qty
        self.cost = cost
        self.order = order


    def __str__(self):
        return f'Sku: {self.sku}, Qty: {self.qty}, Cost: {self.cost}'

    def __repr__(self):
        return f'R: Sku: {self.sku}, Qty: {self.qty}, Cost: {self.cost}'


class BeruOrder:

    def __init__(self, order_id, supplier_order_id, order_date, sku, name, quantity,
                 price, beru_discount, beru_bonuses, status, status_changed_date, payment_type,
                 shipment_wh, region, customer_payment, discount_payment, spasibo_payment, customer_payment_return,
                 discount_payment_return, spasibo_payment_return, claim_payment):
        self.order_id = order_id
        self.supplier_order_id = supplier_order_id
        self.order_date = order_date
        self.sku = sku
        self.name = name
        self.quantity = quantity
        self.price = price
        self.beru_discount = beru_discount
        self.beru_bonuses = beru_bonuses
        self.status = status
        self.status_changed_date = status_changed_date
        self.payment_type = payment_type
        self.shipment_wh = shipment_wh
        self.region = region
        self.customer_payment = customer_payment
        self.discount_payment = discount_payment
        self.spasibo_payment = spasibo_payment
        self.customer_payment_return = customer_payment_return
        self.discount_payment_return = discount_payment_return
        self.spasibo_payment_return = spasibo_payment_return
        self.claim_payment = claim_payment


class BeruAnalizer:

    def __init__(self, filename):
        self.input_file_name = filename
        self._orders = []
        self._bgc = {}
        self._readFile()
        # self._generatebgc()

    def _readFile(self):
        orders = pd.read_excel(self.input_file_name,
                                engine="openpyxl",
                                sheet_name=0,
                                skiprows=1,
                                index_col=None,
                                header=0)

        orders.replace(['nan', 'None', 'NaN'], '')

        for _, row in orders.iterrows():
            self._create_order(row)

    def _create_order(self, row):

        order_id = row['ID заказа']
        supplier_order_id = row['Номер заказа в системе партнера']
        order_date = (row['Дата оформления'])
        sku = row['Ваш SKU']
        name = row['Название товара']
        quantity = row['Количество']
        price = row['Ваша цена\n(за шт.)']
        beru_discount = row['Скидка маркетплейса\n(за шт.)']
        beru_bonuses = row['Оплата бонусами «Спасибо» от Сбербанк\n(за шт.)']
        status = row['Статус заказа']
        status_changed_date = (row['Статус изменён'])
        payment_type = row['Способ оплаты']
        shipment_wh = row['Склад отгрузки']
        region = row['Регион доставки']

        customer_payment = Payment(
            debit=row['Сумма платежа'],
            BGC=row['Номер ПП'],
            BGC_date=row['Дата ПП'],
            payment_id=row['Идентификатор платежа'],
            payment_ladger_date=(row['Дата реестра платежа'])
        )

        discount_payment = Payment(
            debit=row['Сумма платежа.1'],
            BGC=row['Номер ПП.1'],
            BGC_date=row['Дата ПП.1'],
            payment_id=row['Идентификатор платежа.1'],
            payment_ladger_date=(row['Дата реестра платежа.1'])
        )
        spasibo_payment = Payment(
            debit=row['Сумма платежа.2'],
            BGC=row['Номер ПП.2'],
            BGC_date=row['Дата ПП.2'],
            payment_id=row['Идентификатор платежа.2'],
            payment_ladger_date=(row['Дата реестра платежа.2'])
        )
        customer_payment_return = Payment(
            credit=row['Сумма возврата'],
            BGC=row['Номер ПП.3'],
            BGC_date=row['Дата ПП.3'],
            payment_id=row['Идентификатор платежа.3'],
            payment_ladger_date=(row['Дата реестра платежа.3'])
        )
        discount_payment_return = Payment(
            credit=row['Сумма возврата.1'],
            BGC=row['Номер ПП.4'],
            BGC_date=row['Дата ПП.4'],
            payment_id=row['Идентификатор платежа.4'],
            payment_ladger_date=row['Дата реестра платежа.4']
        )
        spasibo_payment_return = Payment(
            credit=row['Сумма возврата.2'],
            BGC=row['Номер ПП.5'],
            BGC_date=row['Дата ПП.5'],
            payment_id=row['Идентификатор платежа.5'],
            payment_ladger_date=row['Дата реестра платежа.5']
        )
        claim_payment = Payment(
            credit=row['Удержанная сумма'],
            BGC=row['Номер ПП.6'],            
            BGC_date=row['Дата ПП.6'],
            payment_id=row['Идентификатор платежа.6'],
            payment_ladger_date=(row['Дата реестра платежа.6'])
        )

        order = BeruOrder(order_id=order_id, supplier_order_id=supplier_order_id, order_date=order_date,
                          sku=sku, name=name, quantity=quantity,
                          price=price, beru_discount=beru_discount, beru_bonuses=beru_bonuses,
                          status=status, status_changed_date=status_changed_date, payment_type=payment_type,
                          shipment_wh=shipment_wh, region=region, customer_payment=customer_payment,
                          discount_payment=discount_payment, spasibo_payment=spasibo_payment,
                          customer_payment_return=customer_payment_return,
                          discount_payment_return=discount_payment_return, 
                          spasibo_payment_return=spasibo_payment_return,
                          claim_payment=claim_payment)

        self._orders.append(order)
        self.checkBGCForOrder(order)

    def checkBGCForOrder(self, order):

        self._checkbgc(order, order.customer_payment)
        self._checkbgc(order, order.discount_payment)
        self._checkbgc(order, order.spasibo_payment)
        self._checkbgc(order, order.customer_payment_return)
        self._checkbgc(order, order.discount_payment_return)
        self._checkbgc(order, order.spasibo_payment_return)
        self._checkbgc(order, order.claim_payment)

    def _checkbgc(self, order, payment):
        print(f'BGS number is {payment.BGC}')
        if not pd.isna(payment.BGC):
            bgc_item = BGCItem(order.sku, order.quantity, payment.payment, order)
            if payment.BGC in self._bgc:
                bgc_item = BGCItem(order.sku, order.quantity, payment.payment, order)
                self._bgc[payment.BGC].addItem(bgc_item)
            else:
                self._bgc[payment.BGC] = BGC(number=payment.BGC, date=payment.BGC_date, item=bgc_item)

    def getBGCExcel(self):
        columns = ['Номер ПП', 'Дата ПП', 'SKU', 'Наименование', 'Количество', 'Стоимость']
        df = pd.DataFrame(columns=columns)

        for number, bgc in self._bgc.items():
            for item in bgc.items:

                data_row = [number, strToDate(bgc.date), item.sku, item.order.name, item.qty, item.cost]
                df.loc[len(df)] = data_row

        now = datetime.now().strftime('%Y %m %d %H-%M')
        file_name = f'beru-payments-{now}.xlsx'

        writer = pd.ExcelWriter(file_name,
                                engine='xlsxwriter',
                                datetime_format='yyyy/mm/dd')
        pt = pd.pivot_table(df, values=['Количество', 'Стоимость'],
                            index=['Дата ПП', 'Номер ПП', 'SKU'],
                            aggfunc={'Количество': np.sum,
                                     'Стоимость': np.sum})
        pt.sort_values('Дата ПП')
        pt = pt.reset_index()
        pt.to_excel(writer, sheet_name='Сводная', index=False)
        pivot_worksheet = writer.sheets['Сводная']
        for i, width in enumerate(get_col_widths(pt)):
            pivot_worksheet.set_column(i-1, i, width)

        df.to_excel(writer, sheet_name='Лист1', index=None, merge_cells=False)
        data_worksheet = writer.sheets['Лист1']

        for i, width in enumerate(get_col_widths(df)):
            data_worksheet.set_column(i-1, i, width)

        writer.save()
