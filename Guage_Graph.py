import os

import PIL
import cx_Oracle
import numpy as np
from datetime import date
from datetime import time
from datetime import datetime
from datetime import timedelta
import plotly.graph_objects as go
import pandas as pd
from IPython.display import Image
from pd2ppt import df_to_powerpoint
from pptx import Presentation
from pptx.util import Inches
from pd2ppt import df_to_table
from pptx.util import Pt
from plotly.subplots import make_subplots
import sys
from PIL import Image
import cv2

class Create_Gauge(object):

    def __init__(
            self,
            max_count
    ):

        self.max_count = max_count
        self.count = 0
        self.columns = 0
        self.rows = 0

    def new_gauge(self, name, stat):

        guage_name = str(name)
        if guage_name == 'SUBTRACT':
            guage_name = '-'
        if int(stat) > int(self.max_count):
            guage_stat = int(self.max_count)
        else:
            guage_stat = int(stat)
        max_count = self.max_count
        red = 'rgb(217, 25, 11)'
        orange = 'rgb(250, 167, 22)'
        green = 'rgb(9, 191, 0)'
        grey = 'rgb(219, 219, 219)'
        colour1 = green
        colour2 = grey


        ratio = guage_stat/max_count
        ratio1 = ratio*75
        ratio2 = 75-ratio1

        if (ratio < 0.01):
            ratio1 = 1
            ratio2 = 75-ratio1

        if ratio >= 0.65:
            colour1 = red
        elif 0.24 < ratio < 0.65:
            colour1 = orange

        fig = make_subplots(rows=1, cols=1, specs=[[{'type':'domain'}]])
        fig.add_trace(go.Pie(
            values=[25, ratio1, ratio2],
            domain={"x": [0, .48]},
            marker_colors=[
                'rgb(255, 255, 255)',
                colour1,
                colour2
            ],
            name="Gauge",
            hole=.7,
            direction="clockwise",
            rotation=135,
            showlegend=False,
            hoverinfo="none",
            sort=False
            # textinfo="label",
            ), 1, 1)

        fig.update_traces(textinfo='none')

        fig.update_layout(
            # Add annotations in the center of the donut pies.
            annotations=[dict(text=guage_name, x=0.5, y=0.53, font_size=40, showarrow=False),
                         dict(text=str(stat), x=0.5, y=0.12, font_size=35, showarrow=False),
                         dict(text='0', x=0.26, y=0.23, font_size=25, showarrow=False),
                         dict(text=self.max_count, x=0.8, y=0.38, font_size=25, showarrow=False)],
            margin=dict(l=30, r=30, t=30, b=0),
        )

        fig.write_image('Images/gauge' + str(self.count) +'.jpg', width=500, height=500,scale=0)
        self.count = self.count+1

    def create_full_image(self, name):

        x = 1
        y = 1
        if self.count > 12:
            self.count = 12
        if self.count < 3:
            if (self.count % 3 == 2):
                self.create_gauge_rows(x, y, 2)
                y = 2
            elif (self.count % 3 == 1):
                # there is only 1 image
                self.create_gauge_rows(x, y, 1)
                y = 2
            x = x+1
        elif 2 < self.count < 10:
            # if self.count > 9:
            while x < (self.count + 1):
                if x % 3 == 0:
                    #there are three images
                    #take the first three images and put them together
                    self.create_gauge_rows(x, y, 3)
                    y = y + 1
                elif (self.count % 3 == 2) and (x > ((self.count // 3) * 3)):
                    self.create_gauge_rows(x, y, 2)
                    y = y + 1
                    break
                elif (self.count % 3 == 1) and (x > ((self.count // 3) * 3)):
                    # there is only 1 image
                    self.create_gauge_rows(x, y, 1)
                    y = y + 1
                    break
                x = x+1
        else:
            while x < (self.count + 1):
                if x % 4 == 0:
                    self.create_four_gauge_rows(x, y, 4)
                    y = y + 1
                elif (self.count % 4 == 3) and (x > ((self.count // 4) * 4)):
                    self.create_four_gauge_rows(x, y, 3)
                    y = y + 1
                    break
                elif (self.count % 4 == 2) and (x > ((self.count // 4) * 4)):
                    self.create_four_gauge_rows(x, y, 2)
                    y = y + 1
                    break
                elif (self.count % 4 == 1) and (x > ((self.count // 4) * 4)):
                    self.create_four_gauge_rows(x, y, 1)
                    y = y + 1
                    break
                x = x + 1
        z = 0
        row_array = list()
        while z < y-1:
            row_array.append(cv2.imread('Images/row' + str(z+1) + '.jpg'))
            z = z + 1
        final = np.concatenate(row_array, axis=0)
        cv2.imwrite('Images/' + str(name) + '.jpg', final)
        self.rows = z

        # filename = 'Images/' + str(name) + '.jpg'
        # W = 500
        # oriimg = cv2.imread(filename)
        # height, width, depth = oriimg.shape
        # imgScale = W / width
        # newX, newY = oriimg.shape[1] * imgScale, oriimg.shape[0] * imgScale
        # newimg = cv2.resize(oriimg, (int(newX), int(newY)), interpolation=cv2.INTER_AREA)
        # cv2.imwrite('Images/' + str(name) + '.jpg', newimg)
        #
        # basewidth = 375
        # img = Image.open('Images/' + str(name) + '.jpg')
        # wpercent = (basewidth / float(img.size[0]))
        # hsize = int((float(img.size[1]) * float(wpercent)))
        # img = img.resize((basewidth, hsize), PIL.Image.ANTIALIAS)
        # img.save('Images/' + str(name) + '.jpg')

    def create_gauge_rows(self, num_gauges, row_num, gauges_per_row):

        if gauges_per_row == 3 and os.path.isfile('Images/gauge' + str(self.count - 3) + '.jpg'):
            img1 = cv2.imread('Images/gauge' + str(num_gauges - 3) + '.jpg')
        # else:
        #     img1 = np.zeros([500, 500, 3], dtype=np.uint8)
        #     img1.fill(255)

        if (gauges_per_row == 3 or gauges_per_row == 2):
            if gauges_per_row == 2:
                img1 = cv2.imread('Images/gauge' + str(self.count - 2) + '.jpg')
            else:
                img2 = cv2.imread('Images/gauge' + str(num_gauges - 2) + '.jpg')
        # else:
        #     img2 = np.zeros([500, 500, 3], dtype=np.uint8)
        #     img2.fill(255)

        if (gauges_per_row == 3 or gauges_per_row == 2 or gauges_per_row == 1):
            if gauges_per_row == 2:
                img2 = cv2.imread('Images/gauge' + str(self.count - 1) + '.jpg')
                img3 = np.zeros([500, 500, 3], dtype=np.uint8)
                img3.fill(255)
            elif gauges_per_row == 1:
                img1 = cv2.imread('Images/gauge' + str(self.count - 1) + '.jpg')
                img2 = np.zeros([500, 500, 3], dtype=np.uint8)
                img2.fill(255)
                img3 = np.zeros([500, 500, 3], dtype=np.uint8)
                img3.fill(255)
            else:
                img3 = cv2.imread('Images/gauge' + str(num_gauges - 1) + '.jpg')
        else:
            img3 = np.zeros([500, 500, 3], dtype=np.uint8)
            img3.fill(255)

        vis = np.concatenate((img1, img2, img3), axis=1)
        cv2.imwrite('Images/row' + str(row_num) + '.jpg', vis)

    def create_four_gauge_rows(self, num_gauges, row_num, gauges_per_row):

        if gauges_per_row == 4 and os.path.isfile('Images/gauge' + str(self.count - 4) + '.jpg'):
            img1 = cv2.imread('Images/gauge' + str(num_gauges - 4) + '.jpg')
        else:
            img1 = np.zeros([500, 500, 3], dtype=np.uint8)
            img1.fill(255)

        if (gauges_per_row == 4 or gauges_per_row == 3)and os.path.isfile('Images/gauge' + str(self.count - 3) + '.jpg'):
            if gauges_per_row == 3:
                img1 = cv2.imread('Images/gauge' + str(self.count - 3) + '.jpg')
            else:
                img2 = cv2.imread('Images/gauge' + str(num_gauges - 3) + '.jpg')
        # else:
        #     img2 = np.zeros([500, 500, 3], dtype=np.uint8)
        #     img2.fill(255)

        if (gauges_per_row == 4 or gauges_per_row == 3 or gauges_per_row == 2) and os.path.isfile(
                'Images/gauge' + str(num_gauges - 2) + '.jpg'):
            if gauges_per_row == 2:
                img1 = cv2.imread('Images/gauge' + str(self.count - 2) + '.jpg')
            elif gauges_per_row == 3:
                img2 = cv2.imread('Images/gauge' + str(self.count - 2) + '.jpg')
            else:
                img3 = cv2.imread('Images/gauge' + str(num_gauges - 2) + '.jpg')
        # else:
        #     img2 = np.zeros([500, 500, 3], dtype=np.uint8)
        #     img2.fill(255)

        if (gauges_per_row == 4 or gauges_per_row == 3 or gauges_per_row == 2 or gauges_per_row == 1) and os.path.isfile(
                'Images/gauge' + str(num_gauges - 1) + '.jpg'):
            if gauges_per_row == 2:
                img2 = cv2.imread('Images/gauge' + str(self.count - 1) + '.jpg')
                img3 = np.zeros([500, 500, 3], dtype=np.uint8)
                img3.fill(255)
                img4 = np.zeros([500, 500, 3], dtype=np.uint8)
                img4.fill(255)
            elif gauges_per_row == 1:
                img1 = cv2.imread('Images/gauge' + str(self.count - 1) + '.jpg')
                img2 = np.zeros([500, 500, 3], dtype=np.uint8)
                img2.fill(255)
                img3 = np.zeros([500, 500, 3], dtype=np.uint8)
                img3.fill(255)
                img4 = np.zeros([500, 500, 3], dtype=np.uint8)
                img4.fill(255)
            elif gauges_per_row == 3:
                img3 = cv2.imread('Images/gauge' + str(self.count - 1) + '.jpg')
                img4 = np.zeros([500, 500, 3], dtype=np.uint8)
                img4.fill(255)
            else:
                img4 = cv2.imread('Images/gauge' + str(num_gauges - 1) + '.jpg')
        # else:
        #     img4 = np.zeros([500, 500, 3], dtype=np.uint8)
        #     img4.fill(255)

        vis = np.concatenate((img1, img2, img3, img4), axis=1)
        cv2.imwrite('Images/row' + str(row_num) + '.jpg', vis)