import os
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


class GaugeChart(object):

    def __init__(
            self,
            presentation,
            df_list,
            name

    ):

        self.presentation = self.presentation
        self.df_list = self.df_list
        self.name = self.name

    def generate_image(self):

        columns = 3
        rows = 3
        num_images = 2

        fig = go.Figure(data=go.Pie(
            values=[50, 10, 10, 10, 10, 10],
            labels=["Log Level \n Test", "Debug", "Info", "Warn", "Error", "Fatal"],
            domain={"x": [0, .48]},
            marker_colors=[
                'rgb(255, 255, 255)',
                'rgb(232,226,202)',
                'rgb(226,210,172)',
                'rgb(223,189,139)',
                'rgb(223,162,103)',
                'rgb(226,126,64)'
            ],
            name="Gauge",
            hole=.3,
            direction="clockwise",
            rotation=90,
            showlegend=False,
            hoverinfo="none",
            textinfo="label",
            textposition="inside"
        ))

        # For numerical labels
        fig.add_trace(go.Pie(
            values=[40, 10, 10, 10, 10, 10, 10],
            labels=["-", "0", "20", "40", "60", "80", "100"],
            domain={"x": [0, .48]},
            marker_colors=['rgba(255, 255, 255, 0)'] * 7,
            hole=.4,
            direction="clockwise",
            rotation=108,
            showlegend=False,
            hoverinfo="none",
            textinfo="label",
            textposition="outside"
        ))

        fig.update_layout(
            xaxis=dict(
                showticklabels=False,
                showgrid=False,
                zeroline=False,
            ),
            yaxis=dict(
                showticklabels=False,
                showgrid=False,
                zeroline=False,
            ),
            shapes=[dict(
                type='path',
                path='M 0.235 0.5 L 0.24 0.65 L 0.245 0.5 Z',
                fillcolor='rgba(44, 160, 101, 0.5)',
                line_width=0.5,
                xref='paper',
                yref='paper')
            ],
            annotations=[
                dict(xref='paper',
                     yref='paper',
                     x=0.23,
                     y=0.45,
                     text='50',
                     showarrow=False
                     )
            ]
        )

        fig.show()
