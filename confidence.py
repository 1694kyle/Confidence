from __future__ import division
import sys
from itertools import izip_longest
import os
from os import path
import shutil
import re
from datetime import datetime

import pandas as pd
import numpy as np
from shapely.geometry import Point, LineString
import openpyxl
from openpyxl import load_workbook
from qt.classes.Dialogs import InputSelectDialog
from PyQt4.QtGui import QApplication

regex_station = re.compile("([\d]+)[+]([\d]{2})[.]([\d]{4})")
round_val = 2


# USER INPUT
def get_survey_file():
    #todo: catch a cancel or invalid survey file
    survey_file = askopenfilename(title='Select Survey File')  # show an "Open" dialog box and return the path to the selected file
    if survey_file == '': sys.exit()
    return survey_file

def input_dialog(projects):
    app = QApplication(sys.argv)
    input_window = InputSelectDialog(projects)
    input_window.show()
    app.exec_()
    survey_file = input_window.survey_path
    survey_file_name = input_window.survey_file_name
    project_name = str(input_window.ui.project_combobox.currentText())
    return survey_file, survey_file_name, project_name


def get_input():
    projects = _get_projects()
    # # todo: move this to network location and populate drop down menu with projects available
    survey_file, survey_file_name, project_name = input_dialog(projects)
    project_path = projects[project_name]

    return survey_file, survey_file_name, project_path


# GET PROJECT FILES
def _get_alignments(project_folder):
    alignment_folder = path.join(project_folder, 'Alignments')
    alignments = [path.join(alignment_folder, alignment) for alignment in os.listdir(alignment_folder)]
    return alignments


def _get_profiles(project_folder):
    profile_folder = path.join(project_folder, 'Profiles')
    profiles = [path.join(profile_folder, profile) for profile in os.listdir(profile_folder)]
    return profiles


# BUILD FRAMES
def build_survey_frame(survey_file):
    headers = ['Point', 'Northing', 'Easting', 'Field Elevation', 'Description']
    frame = pd.io.parsers.read_csv(survey_file, names=headers, header=None)
    frame['Easting'] = frame['Easting'].astype(float)
    frame['Northing'] = frame['Northing'].astype(float)
    frame['Field Elevation'] = frame['Field Elevation'].astype(float)
    # frame.to_excel(r"C:\Users\kbonnet\Desktop\Confidence Shot Tool\survey.xls", index=False)  # for debugging
    return frame

def build_alignment_frame(alignment_file):
    frame = pd.io.excel.read_excel(alignment_file)

    frame['Station'].replace(to_replace=r"[^\d.]+", value='', inplace=True, regex=True)  # convert station to ###.#### format
    frame['Northing'].replace(to_replace=r"[^\d.]+", value='', inplace=True, regex=True)
    frame['Easting'].replace(to_replace=r"[^\d.]+", value='', inplace=True, regex=True)  # convert Northing and Easting column to ####.## format
    frame.drop('Tangential Direction', 1, inplace=True)  # don't need

    return frame.astype(float)

def build_profile_frame(profile_file):
    frame = pd.io.excel.read_excel(profile_file)
    frame['Station'].replace(to_replace=r"[^\d.]+", value='', inplace=True, regex=True)  # convert station column to ####.## format
    frame['Elevation'].replace(to_replace=r"[^\d.]+", value='', inplace=True, regex=True)  # convert elevation column to ####.## format
    frame.drop(['Grade Percent (%)', 'Location'], 1, inplace=True)
    return frame.astype(float)

def build_merged(alignment_files, profile_files):
    alignments = {}
    merged = {}
    for alignment in alignment_files:
        name = path.basename(alignment)[path.basename(alignment).index('[') + 1: path.basename(alignment).index(']')]
        alignments[name] = build_alignment_frame(alignment)

    for profile in profile_files:
        name = path.basename(profile)[path.basename(profile).index('[') + 1:path.basename(profile).index(']')]
        profile_frame = build_profile_frame(profile)
        merged[name] = pd.merge(profile_frame, alignments[name], left_on='Station', right_on='Station')

    # dont use panel because it matches indices for all frames and creates np.nan entries to fil the gaps.
    # when this carries over to creating linestrings, thousands of 0 entries are included and it slows down a lot
    panel = pd.Panel(merged)
    return merged

# CORRELATE SURVEY POINTS WITH ALIGNMENT
def correlate_survey_points(survey_file, alignments, profiles):
    test = False
    merged = build_merged(alignments, profiles)
    survey_frame = build_survey_frame(survey_file)
    e_ls, s_ls = _build_linestrings(merged)
    survey_frame.insert(len(survey_frame.columns), 'Station', np.nan)
    survey_frame.insert(len(survey_frame.columns), 'Offset', np.nan)
    survey_frame.insert(len(survey_frame.columns), 'CL Elevation', np.nan)
    survey_frame.insert(len(survey_frame.columns), 'Crown', -.02)
    # survey_frame.insert(len(survey_frame.columns), 'Nearest Alignment', np.nan): test=True

    for i, row in survey_frame.iterrows():
        point = Point([row.Northing, row.Easting])
        offsets = {}
        for name, ls in e_ls.iteritems():
            offsets[name] = _offset(ls, point)
        min_offset = min(offsets.itervalues())
        nearest_alignment = [k for k in offsets if offsets[k] == min_offset][0]
        # nearest_alignment = min(value for key, value in offsets.iteritems() if value is not None)
        offset = offsets[nearest_alignment]
        nearest_e = _nearest(e_ls[nearest_alignment], point)
        nearest_s = _nearest(s_ls[nearest_alignment], point)
        survey_frame['Offset'][i] = offset
        survey_frame['CL Elevation'][i] = nearest_e.z
        survey_frame['Station'][i] = nearest_s.z
        if test: survey_frame['Nearest Alignment'][i] = nearest_alignment  ### for testing and validation ###

    # convert back to sta format ##+##.##
    survey_frame['Station'] = survey_frame['Station'].round(round_val)
    survey_frame['Station'] = survey_frame['Station'].apply(_convert_to_station)
    return survey_frame


def generate_output_file(frame, survey_file_name):
    headers_round = ['Northing', 'Easting', 'Field Elevation', 'Offset', 'CL Elevation']
    output_folder = r"C:\Users\kbonnet\Desktop\Confidence Shot Tool\Excel Files"
    input_template_file = r"C:\Users\kbonnet\Desktop\Confidence Shot Tool\Template\Confidence Points Template.xlsx"
    file_name = '{}.xlsx'.format(path.splitext(path.basename(survey_file_name))[0])
    output_file = path.join(output_folder, file_name)
    print output_file
    shutil.copy(input_template_file, output_file)

    start_row = 2

    for header in headers_round:
        frame[header] = frame[header].round(round_val)

    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    book = load_workbook(output_file)
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

    frame.to_excel(writer, index=False, startrow=start_row)

    writer.save()


# UTILITY FUNCTIONS
def _build_linestrings(frame_dict):
    e_linestrings = {}
    s_linestrings = {}
    for name, frame in frame_dict.iteritems():
        e_linestrings[name] = LineString([(row.Northing, row.Easting, row.Elevation) for i, row in frame.iterrows()])
        s_linestrings[name] = LineString([(row.Northing, row.Easting, row.Station) for i, row in frame.iterrows()])
    return e_linestrings, s_linestrings


def grouper(n, iterable, fillvalue=None):
    "grouper(3, 'ABCDEFG', 'x') --> ABC DEF Gxx"
    args = [iter(iterable)] * n
    return izip_longest(fillvalue=fillvalue, *args)


def _convert_from_station(x):
    return float(x.replace('+', '').replace(',', ''))


def _convert_to_station(x):
    x = str(x)
    dec_index = x.index('.')
    return '{}+{}'.format(x[:dec_index - 2], x[dec_index - 2:])


def _convert_ne_ele(cols):
    if type(cols) == pd.core.series.Series:
        return cols.applymap(lambda x: float(x.replace(',', '').replace("'", '')))  # ele col only
    else:
        return cols.apply(lambda x: float(x.replace(',', '').replace("'", '')))  # n, e cols


def _convert_angle(col):
    return col.map(lambda x: x.split(u'\xb0')[1].split("'"))


def _nearest(linestring, point):
    # gives the point(x, y, elev) on line that is closest to point
    return linestring.interpolate(linestring.project(point))


def _offset(linestring, point):
    offset = point.distance(_nearest(linestring, point))
    if offset == 0.0:
        return None
    else:
        return offset


def _check_position(a, b, c):
    # a is first point on line, b is second, c is point to be checked
     return ((b.x - a.x)*(c.y - a.y) - (b.y - a.y)*(c.x - a.x)) > 0


def _get_projects():
    project_path = "C:\Users\kbonnet\Desktop\Confidence Shot Tool\Projects"
    projects = {}
    for path, dirs, files in os.walk(project_path):
        if dirs == ['Alignments', 'Profiles']:
            project_name = path[path.index('Projects'):].replace('Projects\\', '')
            projects[project_name] = path
    return projects


def main():
    survey = get_survey_file()
    survey, survey_name, project = get_input()
    start = datetime.now()
    alignments = _get_alignments(project)
    profiles = _get_profiles(project)
    calculated = correlate_survey_points(survey, alignments, profiles)
    generate_output_file(calculated, survey_name)
    diff = datetime.now() - start
    print diff.seconds, 'Second Execution'
    sys.exit()


if __name__ == '__main__':
    main()
# todo: figure out if offset is left or right side of alignment
