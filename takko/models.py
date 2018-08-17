from django.db import models
from django import forms
import pandas as pd
import re
import json
from xlrd import XLRDError


# Create your models here.
class UploadFileForm(forms.Form):
    #title = forms.CharField(max_length=50)
    file = forms.FileField()


class Recipient(object):
    def __init__(self, name, phone_number, address, recipient_orders_df):
        self._name = name
        self._phone_number = phone_number
        self._address = address
        self._recipient_orders_df = recipient_orders_df

        self._zip_code = self.read_zip_code()
        self._old_zip_code = self.read_old_zip_code()
        self._orders = self.read_orders()

        self._combined_order_details = self.combine_order_details()

    @property
    def name(self):
        return self._name

    @property
    def phone_number(self):
        return self._phone_number

    @property
    def address(self):
        return self._address

    @property
    def old_zip_code(self):
        return self._old_zip_code

    @property
    def zip_code(self):
        return self._zip_code

    @property
    def orders(self):
        return self._orders

    def read_orders(self):
        orders = {}
        for order_id in self._recipient_orders_df['주문 번호'].unique():
            mask = self._recipient_orders_df['주문 번호'] == order_id
            orders[int(order_id)] = Order(order_id, self._recipient_orders_df[mask])
        return orders

    def read_zip_code(self):
        zip_codes = self._recipient_orders_df['수취인 우편번호'].unique()
        if len(zip_codes) == 1:
            return zip_codes[0]
        return zip_codes

    def read_old_zip_code(self):
        old_zip_codes = self._recipient_orders_df['수취인 구 우편번호 (6자리)'].unique()
        if len(old_zip_codes) == 1:
            return old_zip_codes[0]
        return old_zip_codes

    @property
    def combined_order_ids(self):
        return {order_id: self._orders[order_id].good_order_ids for order_id in self._orders}

    def combine_order_details(self):
        combined_goods = []
        for order in self._orders.values():
            combined_goods += order.goods

        order_details = {}
        for good in combined_goods:
            if good.name not in order_details.keys():
                order_details[good.name] = {}

            if good.option not in order_details[good.name].keys():
                order_details[good.name][good.option] = good.amount
            else:
                order_details[good.name][good.option] += good.amount
        return order_details

    @property
    def combined_order_details_to_string(self):
        details_str = ''
        for i, good_name in enumerate(self._combined_order_details):
            if i > 0:
                details_str += ' --- '

            details_str += good_name + ': '

            for j, good_option in enumerate(self._combined_order_details[good_name]):
                if j > 0:
                    details_str += ', '

                if not pd.isna(good_option):
                    details_str += good_option + ' '

                good_amount = self._combined_order_details[good_name][good_option]
                if not pd.isna(good_amount):
                    details_str += str(good_amount) + '개'
        return details_str

    @property
    def combined_comments(self):
        comments = pd.Series(self._recipient_orders_df['주문시 남기는 글'].unique()).dropna()
        combined_comments = ''
        for i, comment in enumerate(comments):
            if i > 0:
                combined_comments += ', '
            if len(comment) > 0:
                combined_comments += comment
        return combined_comments


class Order(object):
    def __init__(self, order_id, order_df):
        self._order_id = order_id
        self._order_df = order_df
        self._goods, self._good_order_ids = self.read_goods()
        self._comments = self.read_comments()

    @property
    def order_id(self):
        return self._order_id

    @property
    def comments(self):
        return self._comments

    @property
    def goods(self):
        return self._goods

    @property
    def good_order_ids(self):
        return self._good_order_ids

    def read_goods(self):
        goods = []
        good_order_ids = []
        for _, row in self._order_df.iterrows():
            good_order_id = int(row['상품주문번호'])
            name = row['상품명']
            option = row['옵션정보']
            amount = row['상품수량']
            comment = row['주문시 남기는 글']
            goods.append(Good(good_order_id, name, option, amount, comment))
            good_order_ids.append(good_order_id)
        return goods, good_order_ids

    def read_comments(self):
        return self._order_df['주문시 남기는 글'].unique()


class Good(object):
    def __init__(self, good_order_id, name, option, amount, comment):
        self._good_order_id = good_order_id
        self._name = name
        self._option = option
        self._amount = amount
        self._comment = comment

    @property
    def good_order_id(self):
        return self._good_order_id

    @property
    def name(self):
        return self._name

    @property
    def option(self):
        return self._option

    @property
    def amount(self):
        return self._amount

    @property
    def comment(self):
        return self._comment


class TakkoOrder(object):
    def __init__(self, file_dir):
        self._takko_order_df = self._read_sheet_file(file_dir)
        self._recipients = self._get_unique_recipients()
        self._combined_orders_df = self.combine_all_orders()

    @staticmethod
    def _read_sheet_file(file_dir):
        try:
            df = pd.read_excel(file_dir)
        except XLRDError:
            df = pd.read_html(file_dir, header=0)[0]
        return df

    def _get_unique_recipients(self):
        recipients = []
        recipients_df = self._takko_order_df[['수취인 이름', '수취인 핸드폰 번호', '수취인 전체주소']].drop_duplicates()
        for _, row in recipients_df.iterrows():
            name = row['수취인 이름']
            phone_number = row['수취인 핸드폰 번호']
            address = row['수취인 전체주소']

            mask = _get_mask(self._takko_order_df[['수취인 이름', '수취인 핸드폰 번호', '수취인 전체주소']], row)
            recipients.append(Recipient(name, phone_number, address, self._takko_order_df[mask]))
        return recipients

    def combine_all_orders(self):
        out_columns = ['수취인 이름', '수취인 핸드폰 번호', '수취인 전체주소',
                       '수취인 구 우편번호 (6자리)', '수취인 우편번호',
                       '상품주문번호 리스트', '주문 내역', '주문시 남기는 글']

        combined_orders_df = pd.DataFrame(columns=out_columns)
        for recipient in self._recipients:
            combined_orders_df = combined_orders_df.append(
                {'수취인 이름': recipient.name,
                 '수취인 핸드폰 번호': recipient.phone_number,
                 '수취인 전체주소': recipient.address,
                 '수취인 구 우편번호 (6자리)': recipient.old_zip_code,
                 '수취인 우편번호': recipient.zip_code,
                 '상품주문번호 리스트': json.dumps(recipient.combined_order_ids),
                 '주문 내역': recipient.combined_order_details_to_string,
                 '주문시 남기는 글': recipient.combined_comments}, ignore_index=True)
        return combined_orders_df

    def save_to_excel(self, file_name='combined.xlsx'):
        dfs = {'주문 내역 정리': self._combined_orders_df}
        writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
        for sheetname, df in dfs.items():  # loop through `dict` of dataframes
            df.to_excel(writer, sheet_name=sheetname, index=False)  # send df to writer
            worksheet = writer.sheets[sheetname]  # pull worksheet object

            for idx, col in enumerate(df):  # loop through all columns
                series = df[col]
                max_len = max(
                    series.astype(str).map(visual_len).max(),  # len of largest item
                    visual_len(str(series.name))  # len of column name/header
                )
                worksheet.set_column(idx, idx, max_len)  # set column width

        writer.save()
        return file_name


def _get_mask(dataframe, row):
    return pd.DataFrame(dataframe == row).all(axis='columns')


class TakkoInvoice(object):
    _default_columns = ['번호', '주문번호', '상품주문번호', '배송업체번호', '송장번호', '배송일', '배송완료일']
    _invoice_column_candidates = ['운송장', '운송장번호', '운송장 번호', '송장', '송장번호', '송장 번호']

    def __init__(self, file_dir):
        self._invoice_df = self._read_sheet_file(file_dir)
        if '상품주문번호 리스트' not in self._invoice_df:
            raise Exception('"상품주문번호 리스트" 열이 존재하지 않습니다.')

        has_invoice_column = False
        for column_name in self._invoice_df.columns:
            if column_name in self._invoice_column_candidates:
                has_invoice_column = True
                self._invoice_column = column_name
                break
        if not has_invoice_column:
            raise Exception('운송장번호를 찾을 수 없습니다. 운송장번호를 나타내는 열이 %r 이 중 '
                            '최소 하나의 이름과 일치하여야 합니다.' % self._invoice_column_candidates)

        self._converted_invoice_df = self._convert_invoice_form()

    @staticmethod
    def _read_sheet_file(file_dir):
        try:
            df = pd.read_excel(file_dir)
        except XLRDError:
            df = pd.read_html(file_dir, header=0)[0]
        return df

    @staticmethod
    def _read_combined_order_ids(combined_order_ids):
        combined_order_ids = json.loads(combined_order_ids)
        order_ids_df = pd.DataFrame(columns=['주문번호', '상품주문번호'])
        for order_id in combined_order_ids:
            ids_rows = pd.DataFrame({'주문번호': order_id, '상품주문번호': combined_order_ids[order_id]})
            order_ids_df = order_ids_df.append(ids_rows)
        return order_ids_df

    def _convert_invoice_form(self):
        converted_invoice = pd.DataFrame(columns=self._default_columns)
        count_start = 1
        for _, row in self._invoice_df.iterrows():
            order_ids_df = self._read_combined_order_ids(row['상품주문번호 리스트'])
            count_stop = count_start + len(order_ids_df)
            order_ids_df['번호'] = range(count_start, count_stop)
            order_ids_df['송장번호'] = row[self._invoice_column]
            converted_invoice = converted_invoice.append(order_ids_df)
            count_start = count_stop
        converted_invoice = converted_invoice.reindex(columns=self._default_columns)
        return converted_invoice

    def save_to_excel(self, file_name='invoice.xls'):
        dfs = {'송장 번호 일괄등록': self._converted_invoice_df}
        writer = pd.ExcelWriter(file_name, engine='xlwt')
        for sheetname, df in dfs.items():  # loop through `dict` of dataframes
            df.to_excel(writer, sheet_name=sheetname, index=False)  # send df to writer
            worksheet = writer.sheets[sheetname]  # pull worksheet object

            col_index = 0
            for idx, col in enumerate(df):  # loop through all columns
                series = df[col]
                max_len = max(
                    series.astype(str).map(visual_len).max(),  # len of largest item
                    visual_len(str(series.name))  # len of column name/header
                )
                # worksheet.set_column(idx, idx, max_len)  # set column width
                worksheet.col(col_index).width = int(256 * max_len)
                col_index += 1
        writer.save()
        return file_name


def visual_len(string):
    string_length = len(string)
    charnumeric_length = len(re.findall('\w', string))
    alphanumeric_length = len(re.findall('[A-Za-z0-9]', string))
    korean_length = charnumeric_length - alphanumeric_length
    return string_length + korean_length * 0.75 + 1
