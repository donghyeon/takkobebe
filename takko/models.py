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
        self._name = str(name)
        self._phone_number = str(phone_number)
        self._address = str(address)
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
            orders[order_id] = Order(order_id, self._recipient_orders_df[mask])
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
        return {str(order_id): self._orders[order_id].good_order_ids for order_id in self._orders}

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
                details_str += '; '
            details_str += good_name + ': '
            for j, good_option in enumerate(self._combined_order_details[good_name]):
                if j > 0:
                    details_str += ', '
                if not pd.isna(good_option):
                    details_str += good_option + ' '
                details_str += str(self._combined_order_details[good_name][good_option]) + '개'
        return details_str

    # @property
    # def combined_comments(self):
    #     comments = pd.Series([])
    #     for order in self._orders.values():
    #         comments.append(order.comments)
    #     comments = list(comments.unique().dropna())
    #
    #     combined_comments = ''
    #     for i, comment in enumerate(comments):
    #         if i > 0:
    #             combined_comments += ', '
    #         if len(comment) > 0:
    #             combined_comments += comment
    #     return combined_comments
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
            good_order_id = row['상품주문번호']
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
        self._name = str(name)
        self._option = str(option)
        self._amount = amount
        self._comment = str(comment)

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


class NewTakkoOrder(object):
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


class TakkoOrder(object):
    _recipient_columns = ['수취인 이름', '수취인 핸드폰 번호', '수취인 전체주소']
    _additional_columns = ['수취인 구 우편번호 (6자리)', '수취인 우편번호', '주문시 남기는 글']
    _combined_order_columns = ['상품주문번호 리스트', '주문 내역']
    _order_detail_columns = ['주문번호', '상품주문번호', '상품명', '옵션정보', '상품수량']

    def __init__(self, file_dir):
        self._order_df = self._read_sheet_file(file_dir)
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
        all_recipients = self._order_df[self._recipient_columns]
        return all_recipients.drop_duplicates()

    @property
    def recipients(self):
        return self._recipients

    def combine_all_orders(self):
        all_recipients = self._order_df[self._recipient_columns]
        combined_all_orders = pd.DataFrame(self.recipients,
                                           columns=[*self._recipient_columns,
                                                    *self._combined_order_columns,
                                                    *self._additional_columns])

        for index, recipient in self.recipients.iterrows():
            recipient_mask = self._get_mask(all_recipients, recipient)
            recipient_orders = self._order_df[[*self._order_detail_columns, *self._additional_columns]][recipient_mask]
            combined_all_orders.loc[index, self._combined_order_columns] = self._combine_recipient_orders(recipient_orders)
            combined_all_orders.loc[index, self._additional_columns] = self._combine_additions(recipient_orders)

        self._combined_orders_df = combined_all_orders
        return combined_all_orders

    def _combine_recipient_orders(self, recipient_orders):
        goods = recipient_orders[['상품명']].drop_duplicates()
        combined_recipient_orders = None
        for _, good in goods.iterrows():
            good_mask = self._get_mask(recipient_orders[['상품명']], good)
            good_orders = recipient_orders[good_mask]

            good_details = self._get_good_details(good_orders)
            if combined_recipient_orders is None:
                combined_recipient_orders = good_details
            else:
                combined_recipient_orders += pd.Series(['; ', '; '], index=self._combined_order_columns) + good_details
        return combined_recipient_orders

    def _get_good_details(self, good_orders):
        good_order_numbers = good_orders['상품주문번호'].astype(str)

        good_details = pd.Series(['; '.join(good_order_numbers),
                                  self._get_detail_str(good_orders)], index=self._combined_order_columns)
        return good_details

    def _get_detail_str(self, good_orders):
        detail_str_list = []
        for _, good_row in good_orders.iterrows():
            if not detail_str_list:
                detail_str_list.append(good_row['상품명'] + ': ')
            else:
                detail_str_list.append(', ')

            if not pd.isna(good_row['옵션정보']):
                detail_str_list.append(str(good_row['옵션정보']) + ' ')
            if not pd.isna(good_row['상품수량']):
                detail_str_list.append(str(good_row['상품수량']) + '개')
        return ''.join(detail_str_list)

    def _combine_additions(self, recipient_orders):
        addition_data = pd.Series(index=self._additional_columns)
        for column in self._additional_columns:
            additions_unique = recipient_orders[column].unique()
            additions_str_list = []
            for addition in additions_unique:
                if not pd.isna(addition):
                    additions_str_list.append(str(addition))
            addition_data[column] = '; '.join(additions_str_list)
        return addition_data

    @staticmethod
    def _get_mask(dataframe, row):
        return pd.DataFrame(dataframe == row).all(axis='columns')

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


class TakkoInvoice(object):
    _default_columns = ['번호', '상품주문번호', '배송업체번호', '송장번호', '배송일', '배송완료일']
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

    def _convert_invoice_form(self):
        converted_invoice = pd.DataFrame(columns=self._default_columns)

        p = re.compile('[0-9]+')
        count_start = 1
        for index, row in self._invoice_df.iterrows():
            good_order_numbers = p.findall(row['상품주문번호 리스트'])
            count_stop = count_start + len(good_order_numbers)
            converted_rows = pd.DataFrame({'번호': range(count_start, count_stop),
                                           '상품주문번호': good_order_numbers,
                                           '송장번호': row[self._invoice_column]}, columns=self._default_columns)
            converted_invoice = converted_invoice.append(converted_rows)
            count_start = count_stop

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
