from fpdf import FPDF
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
import matplotlib.ticker as ticker
import datetime
from matplotlib import font_manager

mpl.rcParams['figure.dpi'] = 100


class Report(FPDF):

    def __init__(self, colors, fonts):

        super().__init__()

        self.pdf = FPDF()
        self.pdf.set_auto_page_break(False)

        self.colors = colors

        self.font_families = []
        #         self.mpl_font_paths = []
        self.mpl_font_properties = []
        for font in fonts:
            try:
                self.pdf.add_font(*font)
                self.font_families.append(font[0])

                #                 self.mpl_font_paths.append(Path(mpl.get_data_path(), font[2]))
                self.mpl_font_properties.append(font_manager.FontProperties(fname=font[2]))
            except:
                pass

        self.page_no = 1

    def text_accent(self, x, y, height):

        self.pdf.set_fill_color(*self.colors[0])
        self.pdf.rect(x, y, 1, height, 'F')

    def accented_title(self, x, y, height, font, title, subtitle=None, secondary_font=None):

        self.text_accent(x, y, height)
        self.pdf.set_xy(x + 3, y)

        self.pdf.set_font(*font)

        if not subtitle:
            self.pdf.cell(0, height, title, align='L')

        else:
            self.pdf.cell(0, height / 2, title, align='L')

            self.pdf.set_xy(x + 3, y + height / 2)

            if secondary_font:
                self.pdf.set_font(*secondary_font)

            self.pdf.cell(0, height / 2, subtitle, align='L')

    def add_percentage(self, x, y, width, height, percent, fill_color, fill_color_alt=None, align='C'):

        font_size = self.pdf.font_size_pt
        triangle_width = font_size / 5
        percent_string = f'{np.abs(percent)}%'
        if align == 'C':
            adjust = width / 2 - (triangle_width + 3 * font_size / 10 + self.pdf.get_string_width(percent_string)) / 2
        elif align == 'L':
            adjust = 0
        elif align == 'R':
            adjust = width - (triangle_width + 2 * font_size / 5 + self.pdf.get_string_width(percent_string))

        if percent > 0:

            self.pdf.set_fill_color(*fill_color)

            self.pdf.regular_polygon(x=x + adjust + triangle_width, y=y + height / 2 + triangle_width / 2,
                                     polyWidth=triangle_width, rotateDegrees=270, numSides=3, style='F')

        elif percent < 0:

            if fill_color_alt:
                self.pdf.set_fill_color(*fill_color_alt)

            else:
                self.pdf.set_fill_color(*fill_color)

            self.pdf.regular_polygon(x=x + adjust + triangle_width, y=y + height / 2 + triangle_width / 2,
                                     polyWidth=triangle_width, rotateDegrees=330, numSides=3, style='F')

        self.pdf.set_xy(x + adjust + triangle_width + font_size / 5, y)
        self.pdf.cell(width - adjust - triangle_width - font_size / 5, height, percent_string, align='L', border=0)

    def header(self):
        self.pdf.set_font(self.font_families[1], '', 8)
        self.pdf.set_text_color(100)
        self.pdf.set_xy(10, 0)
        self.pdf.cell(40, 10, 'HAGENBERGSTROM.COM')
        self.pdf.cell(0, 10, 'ANNUAL MARKET REPORT 2022', align='R')
        self.pdf.set_text_color(0)

    def footer(self):
        self.pdf.set_font(self.font_families[2], '', 12)
        self.pdf.set_xy(100, 285)
        self.pdf.cell(10, 5, str(self.page_no), align='C')

    def new_page(self):

        self.pdf.add_page()
        self.header()
        self.footer()
        self.page_no += 1

    def add_cover(self, title, cover_image, logos, subtitle=None):

        self.pdf.add_page()

        self.pdf.image(cover_image, x=0, y=0, w=210)

        self.pdf.set_xy(10, 30)
        self.pdf.set_text_color(255)
        self.pdf.set_font(self.font_families[0], '', 60)
        self.pdf.multi_cell(0, 20, title, align='L', new_x='LMARGIN', new_y='NEXT')
        self.pdf.ln(5)

        if subtitle:
            self.pdf.set_font(self.font_families[1], '', 20)
            self.pdf.multi_cell(0, 15, subtitle, align='L')

        #         self.pdf.set_fill_color(255)
        #         self.pdf.rect(x=150, y=225, w=50, h=62, style='F')
        #         self.pdf.image(logos[0], x=155, y=230, w=40)
        self.pdf.image(logos[0], x=10, y=250, w=60)

        self.pdf.set_text_color(0)

    def add_back_cover(self, logo):

        self.pdf.add_page()

        self.pdf.set_fill_color(*self.colors[0])
        self.pdf.rect(x=0, y=0, w=210, h=297, style='F')

        self.pdf.image(logo, x=75, y=250, w=60)

    def copyright_page(self, logos):

        self.pdf.add_page()

        self.pdf.set_font(self.font_families[2], '', 12)
        self.pdf.set_xy(45, 150)
        self.pdf.multi_cell(120, 5,
                            '©2023 Hagen Bergstrom Team, Coldwell Banker Realty. All rights reserved. All data sourced from Bright MLS. Data deemed reliable but not guaranteed.',
                            align='C')

        self.pdf.image(logos[0], x=160, y=230, w=40)
        self.pdf.image(logos[1], x=10, y=250, w=60)

    def add_table_of_contents(self, regions):

        self.pdf.add_page()

        ownership_types = ['Single Family Residences', 'Condominiums', 'Co-ops']

        self.accented_title(40, 65, 25, (self.font_families[0], '', 48), 'CONTENTS')

        margin = 50
        self.pdf.set_margins(left=margin, top=10, right=margin)

        y = 112.5
        self.pdf.set_xy(margin, 110)
        page_no = 1

        for ownership_type in ownership_types:

            self.pdf.set_font(self.font_families[1], '', 14)
            link = self.pdf.add_link()
            self.pdf.set_link(link, page=page_no + 3)
            self.pdf.cell(50, 5, ownership_type, align='L', link=link)
            self.pdf.cell(0, 5, str(page_no), align='R', new_x='LMARGIN', new_y='NEXT')
            self.pdf.line(margin + 2 + self.pdf.get_string_width(ownership_type), y,
                          210 - margin - (self.pdf.get_string_width(str(page_no)) + 2), y)

            page_no += 1
            y += 10
            self.pdf.ln(5)

            for region in regions:

                if ownership_type in regions[region]['ownership_types']:

                    self.pdf.set_font(self.font_families[1], '', 10)

                    self.pdf.set_margins(left=margin + 5, top=10, right=margin)

                    link = self.pdf.add_link()
                    self.pdf.set_link(link, page=page_no + 3)
                    self.pdf.cell(50, 5, region, align='L', link=link)
                    self.pdf.cell(0, 5, str(page_no), align='R', new_x='LMARGIN', new_y='NEXT')
                    self.pdf.line(7 + margin + self.pdf.get_string_width(region), y,
                                  210 - margin - (self.pdf.get_string_width(str(page_no)) + 2), y)

                    y += 5
                    page_no += 4
                    #                     page_no += 1
                    self.pdf.set_margins(left=margin + 10, top=10, right=margin)

                    if regions[region]['subregions']:
                        all_subregions = regions[region]['subregions']
                        subregions = [subregion for subregion in all_subregions if
                                      (ownership_type in all_subregions[subregion]['ownership_types'])]
                        if len(subregions) >= 3:
                            page_no += ((len(subregions) - 1) // 7 + 1)
                        self.pdf.set_font(self.font_families[1], '', 8)
                        for subregion in subregions:

                            if all_subregions[subregion]['analyze']:
                                link = self.pdf.add_link()
                                self.pdf.set_link(link, page=page_no + 3)
                                self.pdf.cell(50, 5, subregion, align='L', link=link)
                                self.pdf.cell(0, 5, str(page_no), align='R', new_x='LMARGIN', new_y='NEXT')
                                self.pdf.line(12 + margin + self.pdf.get_string_width(subregion), y,
                                              210 - margin - (self.pdf.get_string_width(str(page_no)) + 2), y)
                                page_no += 4
                                #                                 page_no += 1
                                y += 5

            self.pdf.set_margins(left=margin, top=10, right=margin)
            self.pdf.ln(10)
            y += 10

        self.pdf.set_font(self.font_families[1], '', 14)
        link = self.pdf.add_link()
        self.pdf.set_link(link, page=page_no + 2)
        self.pdf.cell(50, 5, '2023 Outlook', align='L', link=link)
        self.pdf.cell(0, 5, str(page_no), align='R', new_x='LMARGIN', new_y='NEXT')
        self.pdf.line(margin + 2 + self.pdf.get_string_width('2023 Outlook'), y,
                      210 - margin - (self.pdf.get_string_width(str(page_no)) + 2), y)

        self.pdf.set_margins(left=10, top=10, right=10)

    def add_table_of_contents_alt(self):

        self.pdf.add_page()

        ownership_types = ['Single Family Residences', 'Condominiums', 'Co-ops']

        self.accented_title(40, 65, 25, (self.font_families[0], '', 48), 'CONTENTS')

        margin = 50
        self.pdf.set_margins(left=margin, top=10, right=margin)

        y = 112.5
        self.pdf.set_xy(margin, 110)
        page_no = 1

        for ownership_type in ownership_types:
            self.pdf.set_font(self.font_families[1], '', 14)
            link = self.pdf.add_link()
            self.pdf.set_link(link, page=page_no + 3)
            self.pdf.cell(50, 5, ownership_type, align='L', link=link)
            self.pdf.cell(0, 5, str(page_no), align='R', new_x='LMARGIN', new_y='NEXT')
            self.pdf.line(margin + 2 + self.pdf.get_string_width(ownership_type), y,
                          210 - margin - (self.pdf.get_string_width(str(page_no)) + 2), y)

            page_no += 5
            y += 15
            self.pdf.ln(10)

        self.pdf.set_margins(left=10, top=10, right=10)

    def section_page(self, image_path, x_right, y, title):

        self.pdf.add_page()

        with self.pdf.local_context(fill_opacity=1):
            self.pdf.image(image_path, x=0, y=0, h=297)

        self.pdf.set_font(self.font_families[0], '', 48)
        self.pdf.set_text_color(255)
        self.pdf.set_xy(10, y)
        self.pdf.multi_cell(x_right - 10, 15, title, align='R')

        self.page_no += 1

    @staticmethod
    def graphic_layout(start_x, start_y, width, height, number_of_items, item_width=None, buffer=None):

        if not buffer:
            buffer = 10

        if number_of_items == 3:

            if not item_width:
                item_width = min((width - buffer) / 2, (height - buffer) / 2)

            coords = [(start_x + item_width / 2 + buffer / 2, start_y),
                      (start_x, start_y + item_width + buffer),
                      (start_x + item_width + buffer, start_y + item_width + buffer)]
            return coords, item_width

        if number_of_items == 4:

            if not item_width:
                item_width = min((width - buffer) / 2, (height - buffer) / 2)

            coords = [(start_x, start_y), (max(start_x + width - item_width, start_x + item_width + buffer), start_y),
                      (start_x, start_y + item_width + buffer),
                      (max(start_x + width - item_width, start_x + item_width + buffer), start_y + item_width + buffer)]
            return coords, item_width

        if number_of_items == 5:

            if not item_width:
                item_width = min((width - buffer) / 2, (height - 2 * buffer) / 3)

            coords = [(start_x, start_y), (max(start_x + width - item_width, start_x + item_width + buffer), start_y),
                      (start_x + width / 2 - item_width / 2, start_y + item_width + buffer),
                      (start_x, start_y + 2 * buffer + 2 * item_width), (
                      max(start_x + width - item_width, start_x + item_width + buffer),
                      start_y + 2 * buffer + 2 * item_width)]

            return coords, item_width

        if number_of_items == 6:

            if not item_width:
                item_width = min((width - buffer) / 2, (height - 2 * buffer) / 3)

            coords = [(start_x, start_y), (max(start_x + item_width + buffer, start_x + width - item_width), start_y),
                      (start_x, start_y + item_width + buffer),
                      (max(start_x + item_width + buffer, start_x + width - item_width), start_y + item_width + buffer),
                      (start_x, start_y + 2 * buffer + 2 * item_width), (
                      max(start_x + item_width + buffer, start_x + width - item_width),
                      start_y + 2 * buffer + 2 * item_width)]

            return coords, item_width

        if number_of_items == 7:

            if not item_width:
                item_width = min((width - 2 * buffer) / 3, (height - 2 * buffer) / 3)

            coords = [(start_x + item_width / 2 + buffer / 2, start_y),
                      (start_x + width - 3 * item_width / 2 - buffer / 2, start_y),
                      (start_x, start_y + item_width + buffer),
                      (start_x + item_width + buffer, start_y + item_width + buffer),
                      (start_x + 2 * item_width + 2 * buffer, start_y + item_width + buffer),
                      (start_x + item_width / 2 + buffer / 2, start_y + 2 * buffer + 2 * item_width),
                      (start_x + width - 3 * item_width / 2 - buffer / 2, start_y + 2 * buffer + 2 * item_width)]

            return coords, item_width

        if number_of_items == 8:
            pass

    def infographic(self, x, y, radius, region_name, data):

        self.pdf.set_fill_color(*self.colors[0])

        x = x + radius
        y = y + radius

        with self.pdf.local_context(fill_opacity=0.3):
            self.pdf.ellipse(x - radius, y - radius, 2 * radius, 2 * radius, 'F')

        radius = radius / 1.25

        self.pdf.set_font(self.font_families[1], '', 0.4 * radius)
        region_width = self.pdf.get_string_width(region_name)
        self.pdf.set_text_color(0)
        self.pdf.set_xy(x - (region_width + radius / 15) / 2, y - radius / 6)
        self.pdf.cell(region_width + radius / 15, radius / 3, region_name)

        self.pdf.set_draw_color(0)
        self.pdf.line(x, y - radius / 5, x, y - radius)
        self.pdf.line(x + (0.866 * radius / 4), y + (0.5 * radius / 4), x + (0.866 * radius), y + (0.5 * radius))
        self.pdf.line(x - (0.866 * radius / 4), y + (0.5 * radius / 4), x - (0.866 * radius), y + (0.5 * radius))

        self.pdf.set_text_color(0)

        self.pdf.set_font(self.font_families[2], '', 0.8 * radius)
        self.pdf.set_xy(x - (0.866 * radius), y - (radius * 5 / 6))
        self.pdf.cell(radius, radius / 3, '{:,}'.format(int(data[0])), align='L')

        self.pdf.set_font('', '', radius / 3)
        self.pdf.set_xy(x - (0.866 * radius), y - (radius / 2))
        self.pdf.cell(radius, radius / 6, 'Homes Sold', align='L')

        self.pdf.set_font('', '', (4 / 15) * radius)
        self.add_percentage(x - (0.866 * radius), y - (radius / 3), radius, radius / 6, data[1], (0, 0, 0), align='L')

        self.pdf.set_font('', '', 0.8 * radius)
        self.pdf.set_xy(x, y - (radius * 5 / 6))
        self.pdf.cell((0.866 * radius), radius / 3, str(int(data[2])), align='R')

        self.pdf.set_font('', '', (4 / 15) * radius)
        self.pdf.set_xy(x, y - (radius / 2))
        self.pdf.multi_cell((0.866 * radius), radius / 12, 'Avg. Days \non Market', align='R')

        self.pdf.set_font('', '', (4 / 15) * radius)
        self.add_percentage(x, y - (radius / 3), 0.866 * radius, radius / 6, data[3], (0, 0, 0), align='R')

        self.pdf.set_font('', '', 0.8 * radius)
        median_sale_price = '${:,}'.format(int(data[4]))
        self.pdf.set_xy(x - self.pdf.get_string_width(median_sale_price) / 2, y + radius / 2)
        self.pdf.cell(self.pdf.get_string_width(median_sale_price), radius / 5, median_sale_price)

        self.pdf.set_font('', '', radius / 3)
        self.pdf.set_xy(x - self.pdf.get_string_width('Median Sale Price') / 2, y + 0.7 * radius)
        self.pdf.cell(self.pdf.get_string_width('Median Sale Price'), radius / 5, 'Median Sale Price')

        self.pdf.set_font('', '', (4 / 15) * radius)
        median_sale_price_yoy = f'{np.abs(data[5])}%'
        self.add_percentage(x - self.pdf.get_string_width(median_sale_price_yoy) / 2, y + 0.9 * radius,
                            self.pdf.get_string_width(median_sale_price_yoy), radius / 6, data[5], (0, 0, 0))

    def infographic_page(self, df, radius, ownership_type, region):

        current_year = 2022

        per_page = 7

        self.new_page()

        self.accented_title(10, 20, 20, (self.font_families[0], '', 24), region['name'].upper(),
                            subtitle='AREA SNAPSHOT', secondary_font=(self.font_families[1], '', 16))
        self.pdf.set_xy(10, 280)
        self.pdf.set_font(self.font_families[2], '', 10)
        self.pdf.multi_cell(80, 4, '*Percentages are year-over-year changes \nfrom 2021 to 2022', align='L')

        all_subregions = region['subregions']
        subregions = [subregion for subregion in all_subregions if
                      (ownership_type in all_subregions[subregion]['ownership_types'])]

        remaining = len(subregions)
        current_index = 0
        while True:

            pages = (remaining - 1) // per_page + 1

            if pages == 1:

                coords, diameter = self.graphic_layout(10, 50, 190, 225, remaining, buffer=5)

            else:

                coords, diameter = self.graphic_layout(10, 50, 190, 225, per_page, buffer=5)

            i = 0
            for subregion in subregions[current_index:current_index + min(remaining, per_page)]:
                metrics = self.generate_metrics(df, ownership=ownership_type, region=all_subregions[subregion])

                self.infographic(*coords[i], diameter / 2, subregion.upper(),
                                 [metrics.loc[str(current_year), 'Sold Listings'],
                                  metrics.loc[str(current_year), 'Sold Listings YoY % Change'],
                                  metrics.loc[str(current_year), 'Sold Average Days on Market'],
                                  metrics.loc[str(current_year), 'Sold Average Days on Market YoY % Change'],
                                  metrics.loc[str(current_year), 'Sold Median Sale Price'],
                                  metrics.loc[str(current_year), 'Sold Median Sale Price YoY % Change']])

                i += 1

            if pages == 1:
                break

            else:
                remaining -= per_page
                current_index += per_page

                self.new_page()
                self.accented_title(10, 20, 20, (self.font_families[0], '', 24), region['name'].upper(),
                                    subtitle="AREA SNAPSHOT (cont'd)", secondary_font=(self.font_families[1], '', 16))

    def table(self, x, y, metrics, status, ownership, cols):

        current_year = 2022
        date = datetime.datetime.strptime(f'{current_year + 1}-01-01', '%Y-%m-%d')

        df = metrics.fillna('N/A')
        df.replace(np.inf, 'N/A', inplace=True)
        df.replace(-1 * np.inf, 'N/A', inplace=True)

        self.pdf.set_xy(x, y)
        self.pdf.set_font(self.font_families[1], '', 14)
        self.pdf.cell(150, 5, f'{status.upper()} LISTINGS', new_x='LMARGIN', new_y='NEXT')
        self.pdf.set_font('', '', 12)
        self.pdf.cell(150, 7, f'{current_year} | {ownership}')

        self.pdf.set_fill_color(50)
        self.pdf.set_text_color(255)
        self.pdf.set_font(self.font_families[2], '', 12)
        self.pdf.set_xy(x, y + 15)
        self.pdf.cell(100, 5, '', fill=True)
        self.pdf.cell(25, 5, str(current_year), fill=True, align='C')
        self.pdf.cell(25, 5, str(current_year - 1), fill=True, align='C')
        self.pdf.cell(40, 5, 'YoY % Change', fill=True, new_x='LMARGIN', new_y='NEXT', align='C')

        table_columns = [[], [], []]
        image_filenames = []
        current_y = y + 20

        self.pdf.set_text_color(0)

        if status == 'Active':
            current_year_index = date
            past_year_index = date.replace(year=date.year - 1)
        else:
            current_year_index = str(current_year)
            past_year_index = str(current_year - 1)

        for col in cols:

            if 'Listings' in col or 'Ratio' in col or 'Supply' in col:

                if status == 'Active':
                    table_columns[0].append(f'{col}*')
                else:
                    table_columns[0].append(col)
            else:
                table_columns[0].append(' '.join(col.split()[1:]))

            if 'Price' in col and 'Ratio' not in col:

                try:
                    table_columns[1].append('${:,}'.format(int(df.loc[current_year_index, col])))
                except:
                    table_columns[1].append('N/A')
                try:
                    table_columns[2].append('${:,}'.format(int(df.loc[past_year_index, col])))
                except:
                    table_columns[2].append('N/A')
            elif 'Ratio' in col:

                try:
                    table_columns[1].append('{}%'.format(int(df.loc[current_year_index, col])))
                except:
                    table_columns[1].append('N/A')

                try:
                    table_columns[2].append('{}%'.format(int(df.loc[past_year_index, col])))
                except:
                    table_columns[2].append('N/A')
            else:

                try:
                    table_columns[1].append(f'{int(df.loc[current_year_index, col])}')
                except:
                    table_columns[1].append('N/A')

                try:
                    table_columns[2].append(f'{int(df.loc[past_year_index, col])}')
                except:
                    table_columns[2].append('N/A')

            self.add_percentage(160, current_y, 40, 6, df.loc[current_year_index, col + ' YoY % Change'], (63, 112, 77),
                                fill_color_alt=(124, 10, 2))

            current_y += 6

        self.pdf.set_xy(x, y + 20)
        self.pdf.multi_cell(100, 6, '\n'.join(table_columns[0]), align='L')
        self.pdf.set_xy(x + 100, y + 20)
        self.pdf.multi_cell(25, 6, '\n'.join(table_columns[1]), align='C')
        self.pdf.set_xy(x + 125, y + 20)
        self.pdf.multi_cell(25, 6, '\n'.join(table_columns[2]), align='C')

        return current_y

    def summary(self, df, place, ownership):

        self.new_page()

        self.accented_title(10, 20, 20, (self.font_families[0], '', 24), place.upper(), subtitle='MARKET SUMMARY',
                            secondary_font=(self.font_families[1], '', 16))

        new_y = self.table(10, 50, df, 'Active', ownership,
                           ['Active Listings', 'Active Average List Price', 'Active Average Days on Market',
                            'Active Median List Price', 'Active Median Days on Market', 'Months of Supply'])
        self.pdf.set_xy(10, new_y + 1)
        self.pdf.set_font(self.font_families[2], '', 10)
        self.pdf.cell(100, 5, '*As of Dec 31, 2022.')
        new_y = self.table(10, new_y + 15, df, 'New', ownership,
                           ['New Listings', 'New Average List Price', 'New Average Days on Market',
                            'New Median List Price', 'New Median Days on Market'])
        new_y = self.table(10, new_y + 10, df, 'Sold', ownership,
                           ['Sold Listings', 'Sold Average List Price', 'Sold Average Sale Price',
                            'Sold/List Price Ratio', 'Sold Average Days on Market', 'Sold Median List Price',
                            'Sold Median Sale Price', 'Sold Median Days on Market'])

    def chart(self, df, bar_vars, line_vars, ylabels, filename):

        colors = ['black', 'red', 'blue', 'green', 'orange', 'magenta']

        fig, ax = plt.subplots(figsize=(12, 6))

        if len(ylabels) == 2:
            ax2 = ax.twinx()

        if len(bar_vars) == 1:
            ax.bar(df.index, df[bar_vars[0][0]],
                   color=(self.colors[0][0] / 255, self.colors[0][1] / 255, self.colors[0][2] / 255), alpha=0.4,
                   width=20, label=bar_vars[0][1])

        elif len(bar_vars) == 2:
            ax.bar(df.index, df[bar_vars[0][0]],
                   color=(self.colors[0][0] / 255, self.colors[0][1] / 255, self.colors[0][2] / 255), alpha=0.4,
                   width=-10, label=bar_vars[0][1], align='edge')
            ax.bar(df.index, df[bar_vars[1][0]],
                   color=(self.colors[2][0] / 255, self.colors[2][1] / 255, self.colors[2][2] / 255), alpha=0.4,
                   width=10, label=bar_vars[1][1], align='edge')

        ax.xaxis.set_major_formatter(mpl.dates.DateFormatter("%b-%y"))
        for label in ax.xaxis.get_ticklabels():
            label.set_fontproperties(self.mpl_font_properties[2])

        for i in range(len(line_vars)):

            if len(ylabels) == 2:
                line_axis = ax2
            else:
                line_axis = ax

            line_axis.plot(df.index, df[line_vars[i][0]], color=colors[i], alpha=0.4, label=line_vars[i][1])

        ax.set_ylabel(ylabels[0], fontproperties=self.mpl_font_properties[2])
        if ylabels[0] == 'Price':
            ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda y, pos: '{:,.0f}'.format(y / 1000) + 'K'))

        for label in ax.yaxis.get_ticklabels():
            label.set_fontproperties(self.mpl_font_properties[2])

        if len(ylabels) == 2:
            ax2.set_ylabel(ylabels[1], rotation=270, labelpad=15, fontproperties=self.mpl_font_properties[2])

            if ylabels[1] == 'Price':
                ax2.yaxis.set_major_formatter(ticker.FuncFormatter(lambda y, pos: '{:,.0f}'.format(y / 1000) + 'K'))

            for label in ax2.yaxis.get_ticklabels():
                label.set_fontproperties(self.mpl_font_properties[2])

        plt.tick_params(left=False, bottom=False)
        plt.grid(color='grey', linestyle='-', linewidth=0.25, alpha=0.4)
        for spine in ['left', 'top', 'right', 'bottom']:
            ax.spines[spine].set_visible(False)

        fig.legend(ncol=len(bar_vars + line_vars), bbox_to_anchor=(0.5, 0), loc='lower center', frameon=False,
                   prop=self.mpl_font_properties[2])

        fig.savefig(filename, bbox_inches='tight')
        plt.close()

    def charts(self, df, base_filename):

        df = df[2:]
        df.set_index((pd.Series(df.index) - pd.DateOffset(months=1)).values, inplace=True)

        self.chart(df, [('New Listings', 'New Listings'), ('Sold Listings', 'Sold Listings')],
                   [('Active Listings', 'Active Listings')], ['Units'],
                   f'{base_filename} {"Active Listings, New Listings, and Sales Per Month"}.png')
        self.chart(df, [('Sold Listings', 'Number of Sales')], [('Sold Median Sale Price', 'Median Sale Price')],
                   ['Units', 'Price'], f'{base_filename} {"Median Sale Price and Number of Sales"}.png')
        self.chart(df, [('Active Median Days on Market', 'Median Days on Market')],
                   [('Sold Median Sale Price', 'Median Sale Price')], ['Days', 'Price'],
                   f'{base_filename} {"Median Sale Price and Median Days on Market"}.png')
        self.chart(df, [],
                   [('< $500k', '< $500k'), ('\$500k - \$750k', '\$500k - \$750k'), ('\$750k - \$1M', '\$750k - \$1M'),
                    ('\$1M - \$1.5M', '\$1M - \$1.5M'), ('\$1.5M - \$2M', '\$1.5M - \$2M'), ('> \$2M', '> \$2M')],
                   ['Days'], f'{base_filename} {"Average Days on Market by Price Range"}.png')
        self.chart(df, [('Active Average List Price', 'Average List Price'),
                        ('Sold Average Sale Price', 'Average Sale Price')],
                   [('Sold/List Price Ratio', 'Sold/List Price Ratio')], ['Price', 'Ratio'],
                   f'{base_filename} {"Average Listing Price and Average Sale Price"}.png')
        self.chart(df, [], [('Months of Supply', 'Months of Supply')], ['Months'],
                   f'{base_filename} {"Months of Supply"}.png')

    def graphic(self, x, y, image_path, title, property_types, descriptions):

        self.pdf.set_xy(x, y)
        self.pdf.set_font(self.font_families[1], '', 16)
        self.pdf.cell(150, 5, title.upper(), new_x='LMARGIN', new_y='NEXT')
        self.pdf.set_font('', '', 14)
        self.pdf.cell(150, 7, '2022 | {}'.format(property_types), new_x='LMARGIN', new_y='NEXT')
        current_y = y + 12
        for variable in descriptions:
            self.pdf.set_font(self.font_families[1], '', 10)
            cell_length = round(self.pdf.get_string_width(variable), 2)
            self.pdf.cell(cell_length + 1, 4, variable)
            self.pdf.set_font(self.font_families[2], '', 10)
            self.pdf.cell(150, 4, '| {}'.format(descriptions[variable]), new_x='LMARGIN', new_y='NEXT')
            current_y += 4

        self.pdf.image(image_path, x=x, y=current_y + 5, h=90)

        return current_y + 90

    def section(self, df, ownership=None, region=None, charts=True):

        metrics = self.generate_metrics(df, ownership=ownership, region=region)

        all_subregions = region['subregions']
        if all_subregions:
            subregions = [subregion for subregion in all_subregions if
                          (ownership in all_subregions[subregion]['ownership_types'])]
            if len(subregions) >= 3:
                self.infographic_page(df, 30, ownership, region)

        self.summary(metrics, region['name'], ownership)

        if charts:
            base_filename = f'{region["name"]} {ownership}'
            print(base_filename)
            self.charts(metrics, base_filename)

            self.new_page()
            self.accented_title(10, 10, 15, (self.font_families[0], '', 20), region['name'])
            new_y = self.graphic(10, 35, f'{base_filename} {"Active Listings, New Listings, and Sales Per Month"}.png',
                                 'Active Listings, New Listings, and Sales Per Month', ownership,
                                 {'Active Listings': 'Number of properties listed for sale at the end of month.',
                                  'New Listings': 'Number of properties newly listed during the month.',
                                  'Sold Properties': 'Number of properties sold during the month.'})
            self.graphic(10, new_y + 10, f'{base_filename} {"Median Sale Price and Number of Sales"}.png',
                         'Median Sale Price and Number of Sales', ownership,
                         {'Median Sale Price': 'Median of sale prices for properties sold during the month.',
                          'Number of Sales': 'Number of properties sold during the month.'})

            self.new_page()
            self.accented_title(10, 10, 15, (self.font_families[0], '', 20), region['name'])
            new_y = self.graphic(10, 35, f'{base_filename} {"Median Sale Price and Median Days on Market"}.png',
                                 'Median Sale Price and Median Days on Market', ownership,
                                 {'Median Sale Price': 'Median of sale prices for properties sold during the month.',
                                  'Median Days on Market': 'Median of days spent on market for all active properties at the end of the month.'})
            self.graphic(10, new_y + 10, f'{base_filename} {"Average Days on Market by Price Range"}.png',
                         'Average Days on Market by Price Range', ownership, {
                             'Average Days on Market': 'Average days spent on market for all active properties at the end of the month.',
                             'Price Range': 'Range of listed price.'})

            self.new_page()
            self.accented_title(10, 10, 15, (self.font_families[0], '', 20), region['name'])
            new_y = self.graphic(10, 35, f'{base_filename} {"Average Listing Price and Average Sale Price"}.png',
                                 'Average Listing Price and Average Sale Price', ownership, {
                                     'Average Listing Price': 'Average list price for all active properties at the end of the month',
                                     'Average Sale Price': 'Average sale price for properties during the month',
                                     'Sold/List Ratio': 'Ratio of sale price to list price for properties sold during the month.'})
            self.graphic(10, new_y + 10, f'{base_filename} {"Months of Supply"}.png', 'Months of Supply', ownership, {
                'Months of Supply': 'Number of months the current inventory will last, given current absorption rate'})

    @staticmethod
    def parse_data(df, start, end, ownership=None, region=None):

        price_thresholds = [0, 500000, 750000, 1000000, 1500000, 2000000, 10000000]

        if ownership:
            df = df[df.Ownership.isin([ownership]).fillna(False)]

        if region:
            df = df[df[region['region_type']].isin(region['labels']).fillna(False)]

        active_listings = df[(df.ListDate < end) & ~(df.OffMarketDate < end) & ~(df.SettledDate < end) & ~(
                    df['Agreement of Sale/Signed Lease Date'] < end)]
        new_listings = df[(df.ListDate >= start) & (df.ListDate < end)]
        sold_listings = df[(df.SettledDate >= start) & (df.SettledDate < end) & (df.Status.isin(['Closed']))]

        row = pd.concat([active_listings['List Price'].describe().round(0).iloc[[0, 1, 5]].set_axis(
            ['Active Listings', 'Active Average List Price', 'Active Median List Price']),
                         active_listings['DOM'].describe().round(0).iloc[[1, 5]].set_axis(
                             ['Active Average Days on Market', 'Active Median Days on Market']),
                         new_listings['List Price'].describe().round(0).iloc[[0, 1, 5]].set_axis(
                             ['New Listings', 'New Average List Price', 'New Median List Price']),
                         new_listings['DOM'].describe().round(0).iloc[[1, 5]].set_axis(
                             ['New Average Days on Market', 'New Median Days on Market']),
                         sold_listings['List Price'].describe().round(0).iloc[[0, 1, 5]].set_axis(
                             ['Sold Listings', 'Sold Average List Price', 'Sold Median List Price']),
                         sold_listings['SoldPrice'].describe().round(0).iloc[[1, 5]].set_axis(
                             ['Sold Average Sale Price', 'Sold Median Sale Price']),
                         sold_listings['DOM'].describe().round(0).iloc[[1, 5]].set_axis(
                             ['Sold Average Days on Market', 'Sold Median Days on Market']),
                         ])

        price_ranges = pd.cut(active_listings['List Price'], price_thresholds, right='False')
        dom_breakdown = active_listings.groupby(by=price_ranges)['DOM'].mean().set_axis(
            ['< $500k', '\$500k - \$750k', '\$750k - \$1M', '\$1M - \$1.5M', '\$1.5M - \$2M', '> \$2M'])
        dom_breakdown.index.names = [None]
        row = pd.concat([row, dom_breakdown])

        row.loc['Sold/List Price Ratio'] = round((sold_listings.SoldPrice / sold_listings['List Price']).mean() * 100,
                                                 2)

        return row

    def generate_metrics(self, df, ownership=None, region=None):

        current_year = 2022

        metrics = pd.DataFrame()

        dates = ['2021-01-01', '2021-02-01', '2021-03-01', '2021-04-01', '2021-05-01', '2021-06-01', '2021-07-01',
                 '2021-08-01', '2021-09-01', '2021-10-01', '2021-11-01', '2021-12-01', '2022-01-01', '2022-02-01',
                 '2022-03-01', '2022-04-01', '2022-05-01', '2022-06-01', '2022-07-01', '2022-08-01', '2022-09-01',
                 '2022-10-01', '2022-11-01', '2022-12-01', '2023-01-01']
        for i in range(len(dates) - 1):
            start = datetime.datetime.strptime(dates[i], '%Y-%m-%d')
            end = datetime.datetime.strptime(dates[i + 1], '%Y-%m-%d')

            metrics[end] = self.parse_data(df, start, end, ownership, region)

        metrics = metrics.T

        metrics['Months of Supply'] = (
                    3 * metrics['Active Listings'] / metrics['Sold Listings'].rolling(window=3).sum()).round(1)

        start = datetime.datetime.strptime('2022-01-01', '%Y-%m-%d')
        end = datetime.datetime.strptime('2023-01-01', '%Y-%m-%d')
        past_start = start.replace(year=current_year - 1)
        past_end = end.replace(year=current_year)

        current_year_metrics = self.parse_data(df, start, end, ownership, region)
        past_year_metrics = self.parse_data(df, past_start, past_end, ownership, region)
        current_year_metrics.name = f'{current_year}'
        past_year_metrics.name = f'{current_year - 1}'

        metrics = pd.concat([past_year_metrics.to_frame().T, metrics], ignore_index=False)
        metrics = pd.concat([current_year_metrics.to_frame().T, metrics], ignore_index=False)

        metrics.loc[f'{current_year}', 'Months of Supply'] = metrics.loc[end, 'Months of Supply']
        metrics.loc[f'{current_year - 1}', 'Months of Supply'] = metrics.loc[past_end, 'Months of Supply']

        metrics_all = metrics.copy()

        for date in metrics.index:

            try:

                yoy = ((metrics.loc[date] / metrics.loc[date.replace(year=date.year - 1)] - 1) * 100).round(1).set_axis(
                    [f'{item} YoY % Change' for item in metrics.columns])

                metrics_all.loc[date, yoy.index] = yoy

            except:

                try:
                    yoy = ((metrics.loc[f'{current_year}'] / metrics.loc[f'{current_year - 1}'] - 1) * 100).round(
                        1).set_axis([f'{item} YoY % Change' for item in metrics.columns])

                    metrics_all.loc[f'{current_year}', yoy.index] = yoy

                except:
                    pass

        #         metrics_all.fillna(0, inplace=True)

        metrics_all.replace(np.inf, np.nan, inplace=True)
        metrics_all.replace(-np.inf, np.nan, inplace=True)
        metrics_all.to_csv(r'C:\Users\Riley Chabot\Downloads\{} {} metrics.csv'.format(region['name'], ownership))
        return metrics_all

    def text_box(self, x, y, text, align='J'):

        self.pdf.set_xy(x, y)
        font_size = self.pdf.font_size
        self.pdf.multi_cell(0, 3 * font_size / 2, text, align=align)

    def forecast(self, sections):

        self.new_page()
        self.accented_title(10, 30, 20, (self.font_families[0], '', 48), '2023 OUTLOOK')
        self.pdf.set_xy(10, 120)

        for i in range(len(sections)):

            self.pdf.set_font(self.font_families[1], '', 20)
            self.text_box(10, self.pdf.get_y() + 10, sections[i][0], align='L')
            self.pdf.ln(5)
            self.pdf.set_font(self.font_families[2], '', 12)
            self.text_box(10, self.pdf.get_y(), sections[i][1])

            if (self.pdf.get_y() > 200) and (i + 1 < len(sections)):
                self.new_page()
                self.pdf.set_xy(10, 20)

    def compose_report(self, df, ownership_types, regions, forecast_text=None, charts=True, output_filename=None):

        self.add_cover('ANNUAL\nMARKET\nREPORT\n2022', r"C:\Users\Riley Chabot\Downloads\IMG_7111.jpg",
                       [r"C:\Users\Riley Chabot\Downloads\HB-Logo-Horizontal_large.png"])
        self.copyright_page([r"C:\Users\Riley Chabot\Downloads\Best Logo.png",
                             r"C:\Users\Riley Chabot\Downloads\HB-Logo-Horizontal_large (1).png"])

        self.add_table_of_contents(regions)

        section_page_params = [(r"C:\Users\Riley Chabot\Downloads\sfr.jpg", 185, 200, 'Single Family\nResidences'),
                               (r"C:\Users\Riley Chabot\Downloads\condo.jpg", 200, 80, 'Condominiums'),
                               (r"C:\Users\Riley Chabot\Downloads\coop.jpg", 190, 255, 'Co-ops')]

        for i in range(len(ownership_types)):

            self.section_page(*section_page_params[i])

            for region in regions:

                if ownership_types[i] in regions[region]['ownership_types']:

                    self.section(df, ownership=ownership_types[i], region=regions[region], charts=charts)

                    subregions = regions[region]['subregions']
                    for subregion in subregions:
                        if subregions[subregion]['analyze'] and (
                                ownership_types[i] in subregions[subregion]['ownership_types']):
                            self.section(df, ownership=ownership_types[i], region=subregions[subregion], charts=charts)

        if forecast_text:
            self.forecast(forecast_text)

        self.add_back_cover(r"C:\Users\Riley Chabot\Downloads\HB-Logo-Horizontal_large.png")

        if output_filename:
            self.pdf.output(output_filename)
