import os
import sys
from datetime import datetime, timedelta
from calendar import monthrange
import json

import pandas as pd
import sqlite3
from ortools.sat.python import cp_model

from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtWidgets import QMessageBox, QComboBox
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QAbstractItemView, QItemDelegate
from PyQt5.QtCore import Qt
from PyQt5.QtCore import QAbstractTableModel
from PyQt5.QtCore import QThread
from PyQt5.QtCore import QObject, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QBrush, QColor, QFont

from scheduler_ui import Ui_MainWindow


connection = sqlite3.connect('db.sqlite3')
cursor = connection.cursor()
cursor.execute("""CREATE TABLE IF NOT EXISTS staffs(
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               staffId INTEGER,
               name TEXT NOT NULL,
               preference TEXT NOT NULL);""")
cursor.execute("""CREATE TABLE IF NOT EXISTS leaders(
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               leaderId INTEGER,
               name TEXT NOT NULL);""")
cursor.execute("""CREATE TABLE IF NOT EXISTS requests(
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               year INTEGER NOT NULL,
               month INTEGER NOT NULL,
               data TEXT NOT NULL);""")
cursor.execute("""CREATE TABLE IF NOT EXISTS schedules(
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               year INTEGER NOT NULL,
               month INTEGER NOT NULL,
               data TEXT NOT NULL);""")


shift_types = ['大夜 (PH)',
               '白班 (1~4)',
               '小夜 (4N)']

day_off1 = 'WW'
day_off2 = 'FF'
business_travel = 'SS'
night_shift = 'PH'
day_shift = 'DAY'
evening_shift = '4N'
staff_required = 9

night_shift_count = 2
evening_shift_count = 3
day_shift_count = 4

unavailable = [day_off1, day_off2, business_travel]

color1 = QBrush(QColor(233, 114, 106))  # E9726A
color2 = QBrush(QColor(106, 170, 233))  # 6AAAE9
color3 = QBrush(QColor(237, 230, 115))  # EDE673
color4 = QBrush(QColor(255, 32, 32))
color5 = QBrush(QColor(110, 216, 96))
#  color5 = QBrush(QColor(16, 45, 162))


def close_db():
  connection.commit()
  connection.close()


class StaffModel(QAbstractTableModel):
  status_message_signal = pyqtSignal(str)

  def __init__(self, parent):
    super(StaffModel, self).__init__(parent)

    self.load_data()

  def load_data(self):
    cursor.execute("""SELECT * FROM staffs;""")
    data = cursor.fetchall()
    self.model_data = [list(row) for row in data]

  def rowCount(self, parent):
    return len(self.model_data)

  def columnCount(self, parent):
    return 4

  def data(self, index, role):
    if role == Qt.DisplayRole:
      return self.model_data[index.row()][index.column()]
    elif role == Qt.EditRole:
      return self.model_data[index.row()][index.column()]
    elif role == Qt.BackgroundRole:
      if index.column() == 2:
        if self.model_data[index.row()][-1] == shift_types[0]:
          return color1
        elif self.model_data[index.row()][-1] == shift_types[2]:
          return color2
    return None

  def setData(self, index, value, role):
    if role == Qt.EditRole:
      self.model_data[index.row()][index.column()] = value

      id = self.model_data[index.row()][0]
      staff_id = str(self.model_data[index.row()][1])
      name = self.model_data[index.row()][2]

      if len(staff_id) == 0:
        try:
          cursor.execute('''UPDATE staffs SET name = "%s" WHERE id = %d''' %
                         (name, id))
          connection.commit()
        except Exception as e:
          self.status_message_signal.emit(
            'Fail to edit staff %d: %s' % (id, str(e)))
      else:
        try:
          staff_id = int(self.model_data[index.row()][1])

          cursor.execute('''UPDATE staffs
                         SET staffId = %d, name = "%s" WHERE id = %d''' %
                         (staff_id, name, id))
          connection.commit()
          return True
        except Exception as e:
          self.status_message_signal.emit(
            'Fail to edit staff %d: %s' % (id, str(e)))
    return False

  def headerData(self, col, orientation, role):
    header = ['ID', 'Staff ID', 'Name', 'Preference']
    if role == Qt.DisplayRole and orientation == Qt.Horizontal:
      return header[col]
    return None

  def flags(self, index):
    if index.column() > 0:
      return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable
    return Qt.ItemIsEnabled | Qt.ItemIsSelectable

  def update(self):
    self.beginResetModel()
    self.load_data()
    self.endResetModel()

  def set_preference(self, index, current_index):
    self.model_data[index.row()][index.column()] = \
        shift_types[current_index]

    id = self.model_data[index.row()][0]
    pref = self.model_data[index.row()][3]
    try:
      cursor.execute('''UPDATE staffs SET preference = "%s" WHERE id = %d''' %
                      (pref, id))
      connection.commit()
    except Exception as e:
      self.status_message_signal.emit(
        'Fail to edit staff %d: %s' % (id, str(e)))

  def add_staff(self, staff_id, name, preference_index):
    try:
      cursor.execute("""INSERT INTO staffs(staffId, name, preference)
                      VALUES(?, ?, ?)""",
                     (staff_id, name, shift_types[preference_index]))
      connection.commit()

      self.update()
    except Exception as e:
      self.status_message_signal.emit(
        'Fail to add staff %d, %s, %d: ' %
        (staff_id, name, preference_index, str(e)))

  def delete_staff(self, selection):
    try:
      for index in selection:
        id = self.model_data[index.row()][0]
        cursor.execute('''DELETE FROM staffs WHERE id = %d''' % (id))
      connection.commit()

      self.update()
    except Exception as e:
      ids = [self.model_data[index.row()][0] for index in selection]
      self.status_message_signal.emit(
        'Fail to delete staff %s: %s' % (str(ids), str(e)))


class StaffItemDelegate(QItemDelegate):
  def __init__(self, parent):
    super(QItemDelegate, self).__init__(parent)

  def createEditor(self, parent, option, index):
    if index.column() == 3:
      combobox = QComboBox(parent)
      combobox.addItems(shift_types)
      return combobox
    return QItemDelegate.createEditor(self, parent, option, index)

  def setEditorData(self, editor, index):
    if index.column() == 3:
      current_index = shift_types.index(index.data())
      editor.setCurrentIndex(current_index)
    else:
      QItemDelegate.setEditorData(self, editor, index)

  def setModelData(self, editor, model, index):
    if index.column() == 3:
      current_index = editor.currentIndex()
      model.set_preference(index, current_index)
    else:
      QItemDelegate.setModelData(self, editor, model, index)


class LeaderModel(QAbstractTableModel):
  status_message_signal = pyqtSignal(str)

  def __init__(self, parent):
    super(LeaderModel, self).__init__(parent)

    self.load_data()

  def load_data(self):
    cursor.execute("""SELECT * FROM leaders;""")
    data = cursor.fetchall()
    self.model_data = [list(row) for row in data]

  def rowCount(self, parent):
    return len(self.model_data)

  def columnCount(self, parent):
    return 3

  def data(self, index, role):
    if role == Qt.DisplayRole:
      return self.model_data[index.row()][index.column()]
    elif role == Qt.EditRole:
      return self.model_data[index.row()][index.column()]
    return None

  def setData(self, index, value, role):
    if role == Qt.EditRole:
      self.model_data[index.row()][index.column()] = value

      id = self.model_data[index.row()][0]
      leader_id = self.model_data[index.row()][1]
      name = self.model_data[index.row()][2]

      try:
        cursor.execute('''UPDATE leaders SET leaderId = %d, name = "%s" WHERE id = %d''' %
                        (leader_id, name, id))
        connection.commit()

        self.status_message_signal.emit('Data Edited')
      except Exception as e:
        self.status_message_signal.emit(
          'Fail to edit leader %d: %s' % (id, str(e)))
    return False

  def headerData(self, col, orientation, role):
    header = ['ID', 'Leader Id', 'Name']
    if role == Qt.DisplayRole and orientation == Qt.Horizontal:
      return header[col]
    return None

  def flags(self, index):
    if index.column() > 0:
      return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable
    return Qt.ItemIsEnabled | Qt.ItemIsSelectable

  def update(self):
    self.beginResetModel()
    self.load_data()
    self.endResetModel()

  def add_leader(self, leader_id, name):
    try:
      cursor.execute("""INSERT INTO leaders(leaderId, name) VALUES(?, ?)""",
                     (leader_id, name,))

      connection.commit()

      self.update()
    except Exception as e:
      self.status_message_signal.emit(
        'Fail to add leader %s: %s' % (name, str(e)))

  def delete_leader(self, selection):
    try:
      for index in selection:
        id = self.model_data[index.row()][0]
        cursor.execute('''DELETE FROM leaders WHERE id = %d''' % (id))
      connection.commit()

      self.update()
    except Exception as e:
      ids = [self.model_data[index.row()][0] for index in selection]
      self.status_message_signal.emit(
        'Fail to delete leader %s: %s' % (str(ids), str(e)))


def check_enough_staff_per_day(day_schedule, required):
  num = 0
  for s in day_schedule:
    if s not in unavailable:
      num += 1
  return num >= required


def to_df(data, current_date, first_day, days_in_month):
  df_data = {
    'ID': [row[0][1] if i >= 4 else '' for i, row in enumerate(data)],
    'Name': [row[0][2] if i >= 4 else '' for i, row in enumerate(data)],
    'Limit': [row[1] for row in data],
  }
  columns = ['ID', 'Name', 'Limit']

  day_text = ['一', '二', '三', '四', '五', '六', '日']
  for day in range(2, days_in_month+2):
    text = day_text[(first_day + day - 2) % len(day_text)]
    colname = '%d/%d\n%s' % (current_date.month, day-1, text)
    columns.append(colname)
    df_data[colname] = [row[day] for row in data]

  columns.append('Total')
  df_data['Total'] = [row[-1] for row in data]
  df = pd.DataFrame(df_data, columns=columns)
  return df


def from_df(df):
  cursor.execute("""SELECT * FROM staffs;""")
  staffs = cursor.fetchall()

  cursor.execute("""SELECT * FROM leaders;""")
  leaders = cursor.fetchall()

  data = []
  for i, row in enumerate(df.values):
    if i == 0:
      row_data = ['Days off', ''] + [True if item == 'True' else False for item in row[4:-1]] + \
        [int(row[-1])]
    elif i == 1:
      row_data = [night_shift, ''] + [int(item) for item in row[4:-1]] + ['']
    elif i == 2:
      row_data = [day_shift, ''] + [int(item) for item in row[4:-1]] + ['']
    elif i == 3:
      row_data = [evening_shift, ''] + [int(item) for item in row[4:-1]] + ['']
    elif i == 4:
      leader_id = row[1] if not pd.isnull(row[1]) else 0
      leader_name = row[2]

      leader = [0, leader_id, leader_name]

      for l in leaders:
        if l[2] == leader_name:
          leader = [c for c in l]
          break

      row_data = [leader]
      for item in row[3:]:
        if not pd.isnull(item):
          if isinstance(item, str):
            row_data.append(item)
          elif isinstance(item, bool):
            row_data.append(item)
          else:
            row_data.append(int(item))
        else:
          row_data.append('')
    elif i >= 5:
      staff_id = row[1] if not pd.isnull(row[1]) else 0
      staff_name = row[2]

      staff = [0, staff_id, staff_name, shift_types[1]]
      for s in staffs:
        if s[2] == staff_name:
          staff = [c for c in s]
          break

      row_data = [staff]
      for item in row[3:]:
        if not pd.isnull(item):
          if isinstance(item, str):
            row_data.append(item)
          elif isinstance(item, bool):
            row_data.append(item)
          else:
            row_data.append(int(item))
        else:
          row_data.append('')
    data.append(row_data)
  return data


class RequestModel(QAbstractTableModel):
  status_message_signal = pyqtSignal(str)

  def __init__(self, parent, current_date):
    super(RequestModel, self).__init__(parent)

    self.set_current_date(current_date)

  def set_current_date(self, current_date):
    self.current_date = current_date
    self.first_day, self.days_in_month = monthrange(current_date.year,
                                                    current_date.month)
    self.load_data()

  def load_data(self):
    """
    data should look like this

          | limit  | extra?      | extra?      | .... |
    name1 | limit1 | preference1 | preference2 | .... |  total days off |
    name2 | limit2 | preference1 | preference2 | .... |  total days off |
    """

    self.beginResetModel()
    cursor.execute("""SELECT * FROM staffs;""")
    staffs = cursor.fetchall()

    cursor.execute("""SELECT * FROM leaders;""")
    leaders = cursor.fetchall()

    cursor.execute("""SELECT data FROM requests
                   WHERE year = %d and month = %d""" %
                   (self.current_date.year, self.current_date.month))
    data = cursor.fetchall()
    if len(data) > 0:
      self.model_data = json.loads(data[0][0])

      # find leader
      self.leader = [0, 0, 'Unknown']
      if len(leaders) > 0:
        self.model_data[4][0] = leaders[-1]
        self.leader = leaders[-1]
      else:
        self.model_data[4][0] = self.leader

      # refresh staffs
      staff_model_data = []
      for staff in staffs:
        found = False
        for s in range(5, len(self.model_data)):
          if self.model_data[s][0][2] == staff[2]:
            staff_model_data.append(self.model_data[s])
            found = True
        if not found:
          staff_model_data.append([staff, ''] + ['' for _ in range(self.days_in_month)] + [0])

      self.model_data = self.model_data[:5] + staff_model_data
    else:
      self.model_data = []

      # insert scheduled days off
      days_off = ['Days off', '']
      for d in range(self.days_in_month):
        if (d + self.first_day) % 7 in [5, 6]:
          days_off.append(True)
        else:
          days_off.append(False)
      days_off.append(0)
      self.model_data.append(days_off)

      # insert extra staff options
      row_data = [night_shift, ''] + [2 for _ in range(self.days_in_month)] + ['']
      self.model_data.append(row_data)
      row_data = [day_shift, ''] + [4 for _ in range(self.days_in_month)] + ['']
      self.model_data.append(row_data)
      row_data = [evening_shift, ''] + [3 for _ in range(self.days_in_month)] + ['']
      self.model_data.append(row_data)

      # insert leader
      self.leader = None
      if len(leaders) > 0:
        self.leader = leaders[0]
        row_data = [self.leader, ''] + ['' for _ in range(self.days_in_month)] + [0]
      else:
        row_data = [[0, 0, 'Unknown'], ''] + ['' for _ in range(self.days_in_month)] + [0]
      self.model_data.append(row_data)

      # insert staff
      for staff in staffs:
        row_data = [staff, ''] + ['' for _ in range(self.days_in_month)] + [0]
        self.model_data.append(row_data)

    self.save()
    self.update_states()
    self.endResetModel()

  def update_states(self):
    """
    update the states of request model
      - check if each day should have enough staff for work
      - compute the total number day off each staff should have
      - compute the total day off that each staff request
    """

    # each day should have enough staff for work
    self.enough = [[True for _ in range(self.days_in_month+3)]
                         for _ in range(len(self.model_data))]

    for day in range(2, self.days_in_month+2):
      day_schedule = [self.model_data[s][day] for s in range(4, len(self.model_data))]
      total_staff = self.model_data[1][day] + self.model_data[2][day] + \
        self.model_data[3][day]
      if not check_enough_staff_per_day(day_schedule, total_staff):
        for i in range(4, len(self.model_data)):
          self.enough[i][day] = False

    # compute total number of day off
    total_day_off = 0
    for day in range(2, self.days_in_month+2):
      if self.model_data[0][day]:
        total_day_off += 1
    self.model_data[0][-1] = total_day_off

    # compute total requested day off for each staff
    for s in range(4, len(self.model_data)):
      total_day_off = 0
      for day in range(2, self.days_in_month+2):
        if self.model_data[s][day] in unavailable:
          total_day_off += 1
      self.model_data[s][-1] = total_day_off

  def update_shift_count(self, day):
    """
    - if leader has a day off, day shift should have additional staff
    - if leader has evening shift, day shift should also have addition staff
      but evening shift require 1 less staff
    """

    if self.model_data[4][day] in unavailable and not self.model_data[0][day]:
      self.model_data[2][day] = day_shift_count + 1
    elif self.model_data[4][day] == evening_shift:
      if not self.model_data[0][day]:
        self.model_data[2][day] = day_shift_count + 1
      self.model_data[3][day] = evening_shift_count - 1
    else:
      self.model_data[1][day] = night_shift_count
      self.model_data[2][day] = day_shift_count
      self.model_data[3][day] = evening_shift_count

  def save(self):
    try:
      cursor.execute("""SELECT data FROM requests
                    WHERE year = %d and month = %d""" %
                    (self.current_date.year, self.current_date.month))
      data = cursor.fetchall()
      json_data = json.dumps(self.model_data)

      if len(data) > 0:
        cursor.execute("""UPDATE requests set data = '%s'
                       WHERE year = %d and month = %d""" %
                       (json_data, self.current_date.year, self.current_date.month))
      else:
        cursor.execute("""INSERT INTO requests(year, month, data)
                       VALUES(?, ?, ?)""", (self.current_date.year,
                                            self.current_date.month,
                                            json_data))
      connection.commit()
    except Exception as e:
      self.status_message_signal.emit(str(e))

  def set_values(self, selection, value):
    self.beginResetModel()
    updated = False
    for index in selection:
      if index.row() > 4 and \
          index.column() > 1 and index.column() <= self.days_in_month+1:
        self.model_data[index.row()][index.column()] = value
        updated = True

    if updated:
      self.update_states()
      self.save()

    self.endResetModel()
    return updated

  def export_json(self, filepath):
    obj = {
      'year': self.current_date.year,
      'month': self.current_date.month,
      'data': self.model_data,
    }

    try:
      with open(filepath, 'w') as f:
        json.dump(obj, f)
      return True
    except Exception:
      return False

  def export_csv(self, filepath):
    df = to_df(self.model_data, self.current_date,
               self.first_day, self.days_in_month)
    df.to_csv(filepath)

  def export_excel(self, filepath):
    df = to_df(self.model_data, self.current_date,
               self.first_day, self.days_in_month)
    df.to_excel(filepath)

  def import_json(self, filepath):
    with open(filepath) as f:
      obj = json.load(f)
    new_model_data = obj['data']

    # check if size is the same
    if len(new_model_data) != len(self.model_data):
      return False

    for i, row in enumerate(new_model_data):
      if len(row) != len(self.model_data[i]):
        return False

    for s in range(len(self.model_data)):
      for day in range(self.columnCount(0)):
        if not isinstance(self.model_data[s][day], type(new_model_data[s][day])):
          print(self.model_data[s][day], new_model_data[s][day])
          return False

    self.status_message_signal.emit('import success')
    self.beginResetModel()
    self.model_data = new_model_data
    self.endResetModel()
    self.save()
    return True

  def import_df(self, df):
    try:
      new_model_data = from_df(df)

      # check if size is the same
      if len(new_model_data) != len(self.model_data):
        return False

      for i, row in enumerate(new_model_data):
        if len(row) != len(self.model_data[i]):
          return False

      for s in range(len(self.model_data)):
        for day in range(self.columnCount(0)):
          if not isinstance(self.model_data[s][day], type(new_model_data[s][day])):
            return False

      self.status_message_signal.emit('import success')
      self.beginResetModel()
      self.model_data = new_model_data
      self.update_states()
      self.endResetModel()
      self.save()
      return True
    except Exception:
      self.status_message_signal.emit('fail to import data')

  def import_csv(self, filepath):
    df = pd.read_csv(filepath)
    self.import_df(df)

  def import_excel(self, filepath):
    df = pd.read_excel(filepath)
    self.import_df(df)

  def rowCount(self, parent):
    return len(self.model_data)

  def columnCount(self, parent):
    return self.days_in_month + 3

  def data(self, index, role):
    if role == Qt.DisplayRole:
      if index.row() >= 4 and index.column() == 0:
        return self.model_data[index.row()][index.column()][2]
      else:
        return self.model_data[index.row()][index.column()]
    elif role == Qt.EditRole:
      return self.model_data[index.row()][index.column()]
    elif role == Qt.BackgroundColorRole:
      if index.row() == 0:
        if self.model_data[index.row()][index.column()] and index.column() > 0:
          return color4
      elif index.row() == 1:
        return color1
      elif index.row() == 3:
        return color2
      else:
        if index.column() == 0:
          offset = 5
          if index.row() >= offset:
            if self.model_data[index.row()][0][3] == shift_types[0]:
              return color1
            elif self.model_data[index.row()][0][3] == shift_types[2]:
              return color2
        else:
          if self.model_data[index.row()][index.column()] == day_off1:
            return color3
          elif self.model_data[index.row()][index.column()] == day_off2:
            return color3
          elif self.model_data[index.row()][index.column()] == business_travel:
            return color3
          elif self.model_data[index.row()][index.column()] == night_shift:
            return color1
          elif self.model_data[index.row()][index.column()] == evening_shift:
            return color2
    elif role == Qt.ForegroundRole:
      if not self.enough[index.row()][index.column()]:
        return color4
    return None

  def setData(self, index, value, role):
    changed = False
    if role == Qt.EditRole:
      if index.column() > 0 and index.column() <= self.days_in_month:
        if index.row() == 0:
          self.model_data[index.row()][index.column()] = value
          changed = True
        elif index.row() > 0 and index.row() < 4:
          if value >= 0:
            self.model_data[index.row()][index.column()] = value
            changed = True
          else:
            self.status_message_signal.emit('invalid number of staffs')
        elif index.row() >= 4:
          # edit staff day off or shift request
          value = value.upper()
          possible_val = [day_off1, day_off2,
                          night_shift, day_shift, evening_shift,
                          business_travel, '']
          if value in possible_val:
            self.model_data[index.row()][index.column()] = value
            changed = True

          if index.row() == 4 and \
              index.column() >= 1 and index.column() <= self.days_in_month:
            self.update_shift_count(index.column())

    if changed:
      self.update_states()
      self.save()
    return changed

  def headerData(self, col, orientation, role):
    if role == Qt.DisplayRole:
      if orientation == Qt.Horizontal:
        day_text = ['一', '二', '三', '四', '五', '六', '日']
        if col == 0:
          return 'Name'
        elif col == 1:
          return 'Limit'
        elif col > 1 and col <= self.days_in_month+1:
          text = day_text[(self.first_day + col - 2) % len(day_text)]
          return '%d/%d\n%s' % (self.current_date.month, col-1, text)
        else:
          return 'Total'
      elif orientation == Qt.Vertical:
        if col == 0:
          return 'off'
        elif col == 1:
          return night_shift
        elif col == 2:
          return day_shift
        elif col == 3:
          return evening_shift
        elif col == 4:
          return self.model_data[col][0][1]
        elif col >= 5:
          return self.model_data[col][0][1]
    return None

  def flags(self, index):
    if index.column() > 0:
      return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable
    return Qt.ItemIsEnabled | Qt.ItemIsSelectable


class AsyncTask(QThread):
  def __init__(self, parent, func):
    QThread.__init__(self, parent)
    self.func = func

  def run(self):
    self.func()


class ScheduleModel(QAbstractTableModel):
  status_message_signal = pyqtSignal(str)
  set_optimize_status_signal = pyqtSignal(str)

  def __init__(self, parent, current_date):
    super(ScheduleModel, self).__init__(parent)

    self.parent = parent
    self.set_current_date(current_date)
    self.work_day_constrain = 7
    self.day_off_constrain = 1

  def set_work_day_constrain(self, constrain):
    try:
      self.work_day_constrain = int(constrain)
    except Exception:
      pass

  def set_day_off_contrain(self, constrain):
    try:
      self.day_off_constrain = int(constrain)
    except Exception:
      pass

  def set_current_date(self, current_date):
    self.current_date = current_date
    self.first_day, self.days_in_month = monthrange(current_date.year,
                                                    current_date.month)
    self.load_data()

  def load_staffs(self):
    cursor.execute("""SELECT * FROM staffs;""")
    self.staffs = cursor.fetchall()

  def load_request(self):
    cursor.execute("""SELECT data FROM requests
                   WHERE year = %d and month = %d;""" %
                   (self.current_date.year, self.current_date.month))
    data = cursor.fetchall()

    if len(data) > 0:
      self.preference_data = json.loads(data[0][0])
    else:
      self.status_message_signal.emit('Request not set yet')

  def load_data(self):
    """
    load the data if exists from db
    else copy the preference as current model data
    """

    self.beginResetModel()
    self.load_staffs()
    self.load_request()

    cursor.execute("""SELECT * FROM schedules
                   WHERE year = %d and month = %d;""" %
                   (self.current_date.year, self.current_date.month))
    data = cursor.fetchall()
    if len(data) > 0:
      self.schedule_data = json.loads(data[0][3])

      # copy leader's shift
      for day in range(2, self.days_in_month+2):
        if self.preference_data[4][day]:
          self.schedule_data[4][day] = self.preference_data[4][day]
        else:
          self.schedule_data[4][day] = day_shift
      self.schedule_data[4][0] = self.preference_data[4][0]

      # refresh staff members
      staff_model_data = []
      for staff in self.staffs:
        found = False
        for s in range(5, len(self.schedule_data)):
          if self.schedule_data[s][0][2] == staff[2]:
            staff_model_data.append(self.schedule_data[s])
            found = True
        if not found:
          staff_model_data.append([staff, ''] + ['' for _ in range(self.days_in_month)] + [0])

      self.schedule_data = self.schedule_data[:5] + staff_model_data

      # copy the staff requirment and staff required shift type
      for i in range(4):
        for j in range(2, self.columnCount(0)):
          self.schedule_data[i][j] = self.preference_data[i][j]

      for i in range(5, len(self.schedule_data)):
        self.schedule_data[i][1] = self.preference_data[i][1]
    else:
      # direct copy from preference data
      self.schedule_data = []
      for row in self.preference_data:
        r = [col for col in row]
        self.schedule_data.append(r)

    self.highlight()
    self.endResetModel()

  def save(self):
    try:
      cursor.execute("""SELECT data FROM schedules
                    WHERE year = %d and month = %d""" %
                    (self.current_date.year, self.current_date.month))
      data = cursor.fetchall()
      json_data = json.dumps(self.schedule_data)

      if len(data) > 0:
        cursor.execute("""UPDATE schedules set data = '%s'
                       WHERE year = %d and month = %d""" %
                       (json_data, self.current_date.year, self.current_date.month))
      else:
        cursor.execute("""INSERT INTO schedules(year, month, data)
                       VALUES(?, ?, ?)""", (self.current_date.year,
                                            self.current_date.month,
                                            json_data))
      connection.commit()
    except Exception as e:
      self.status_message_signal.emit(str(e))

  def load_previous_month_data(self):
    """
    load the schedule for last month for optimization
    """

    first = datetime(year=self.current_date.year,
                     month=self.current_date.month,
                     day=1)
    last_month = first - timedelta(days=1)

    cursor.execute("""SELECT * FROM schedules
                   WHERE year = %d and month = %d;""" %
                   (last_month.year, last_month.month))
    data = cursor.fetchall()
    if len(data) > 0:
      self.last_month_data = json.loads(data[0][3])

  def prev_month_schedule_last_few_days(self, staffs):
    constrain_days = self.work_day_constrain-1
    prev_data = {s: [0 for _ in range(constrain_days)] for s in staffs}
    if hasattr(self, 'last_month_data'):
      for name in staffs:
        # find the staff
        for s in range(5, len(self.last_month_data)):
          if self.last_month_data[s][0][2] == name:
            row_schedule = [col for col in self.last_month_data[s][-1-constrain_days:-1]]
            cummulative_count = 0
            for j, col in enumerate(row_schedule[::-1]):
              if col not in unavailable:
                cummulative_count += 1
              prev_data[name][-(j+1)] = cummulative_count
    return prev_data

  def optimize_asyn(self):
    self.load_previous_month_data()
    if hasattr(self, 'task') and not self.task.isFinished():
      self.status_message_signal.emit('Optimization is still running...')
    else:
      self.task = AsyncTask(self.parent, self.optimize_sync)
      self.task.setTerminationEnabled(True)
      self.task.finished.connect(self.save)
      self.task.start()

  def optimize_sync(self):
    self.beginResetModel()
    self.optimize()
    self.update_state()
    self.endResetModel()

  def optimize(self):
    """
    optimize using CP-SAT model from google
    """

    self.status_message_signal.emit('start optimization...')
    model = cp_model.CpModel()

    shifts = {}

    # add variables
    num_shifts = 3
    staffs = []
    for s in range(5, len(self.schedule_data)):
      staff = self.schedule_data[s][0][2]
      staffs.append(staff)
      for day in range(2, self.days_in_month+2):
        for n in range(num_shifts):
          shifts[(staff, day, n)] = model.NewBoolVar(
            'shift_staff%s_day%d_shift%d' % (staff, day, n))

    # each day should have the required number of staff for work
    for day in range(2, self.days_in_month+2):
      for n in range(num_shifts):
        total = sum(shifts[(staff, day, n)] for staff in staffs)
        model.Add(total >= self.schedule_data[n+1][day])

    # every staff should only work when there are at least 16 hour
    # in between each shift
    for staff in staffs:
      for day in range(2, self.days_in_month+2):
        model.Add(sum(shifts[(staff, day, n)] for n in range(num_shifts)) <= 1)

        if day < self.days_in_month:
          model.Add(shifts[(staff, day, 1)] + shifts[(staff, day, 2)] + shifts[(staff, day+1, 0)] <= 1)
          model.Add(shifts[(staff, day, 2)] + shifts[(staff, day+1, 0)] + shifts[(staff, day+1, 1)] <= 1)

    preference_shift_count = 16
    for s in range(5, len(self.schedule_data)):
      staff = self.schedule_data[s][0][2]
      staff_pref = self.schedule_data[s][0][3]

      total = sum(shifts[(staff, day, n)]
                  for day in range(2, self.days_in_month+2)
                  for n in range(3))
      model.Add(total <= self.days_in_month)

      # if this staff prefer night shift or evening shift
      # this staff should have at least 16 that kind of shift for this month
      if staff_pref == shift_types[0]:
        total = sum(shifts[(staff, day, 0)] for day in range(2, self.days_in_month+2))
        model.Add(total >= preference_shift_count)
      elif staff_pref == shift_types[2]:
        total = sum(shifts[(staff, day, 2)] for day in range(2, self.days_in_month+2))
        model.Add(total >= preference_shift_count)
        #  model.Maximize(total)

    # for a number days of work, each staff should have at least some days off
    for start_day in range(2, self.days_in_month+2-self.work_day_constrain):
      for staff in staffs:
        total = sum(shifts[(staff, day, n)]
                    for day in range(start_day, start_day+self.work_day_constrain)
                    for n in range(num_shifts))
        model.Add(total <= self.work_day_constrain-self.day_off_constrain)

    # also check the work days from previous month
    prev_month_data = self.prev_month_schedule_last_few_days(staffs)
    for num_day in range(1, self.work_day_constrain):
      for staff in staffs:
        already_working = prev_month_data[staff][num_day-1]
        total = sum(shifts[(staff, day, n)]
                    for day in range(2, 2+num_day)
                    for n in range(num_shifts))
        model.Add(total <= self.work_day_constrain-self.day_off_constrain-already_working)

    # staff should not be working
    # when business travel is scheduled
    for s in range(5, len(self.preference_data)):
      staff = self.schedule_data[s][0][2]
      for day in range(2, self.days_in_month+2):
        if self.preference_data[s][day] == business_travel:
          model.Add(sum(shifts[(staff, day, n)] for n in range(num_shifts)) == 0)

    # every staff should have at least the same number of day off
    # as the number of saturday and sunday plus
    # the number of national holiday
    total_number_day_off_required = self.preference_data[0][-1]
    num_work_days = self.days_in_month-total_number_day_off_required
    for s in range(5, len(self.preference_data)):
      staff = self.schedule_data[s][0][2]
      model.Add(sum(shifts[(staff, day, n)]
                    for day in range(2, self.days_in_month+2)
                    for n in range(num_shifts)) <= num_work_days)

    # if this staff is limited to the type of shift
    # only give the staff that shift
    for s in range(5, len(self.preference_data)):
      limited_shift = self.schedule_data[s][1]
      if limited_shift:
        #  print(staff, self.schedule_data[s][1])
        staff = self.schedule_data[s][0][2]
        if limited_shift == night_shift:
          model.Add(sum(shifts[(staff, day, 1)]
                        for day in range(2, self.days_in_month+2)) == 0)
          model.Add(sum(shifts[(staff, day, 2)]
                        for day in range(2, self.days_in_month+2)) == 0)
        elif limited_shift == day_shift:
          model.Add(sum(shifts[(staff, day, 0)]
                        for day in range(2, self.days_in_month+2)) == 0)
          model.Add(sum(shifts[(staff, day, 2)]
                        for day in range(2, self.days_in_month+2)) == 0)
        elif limited_shift == evening_shift:
          model.Add(sum(shifts[(staff, day, 0)]
                        for day in range(2, self.days_in_month+2)) == 0)
          model.Add(sum(shifts[(staff, day, 1)]
                        for day in range(2, self.days_in_month+2)) == 0)

    # optimization objective
    # try to satisfy the request made by staffs
    objective = 0
    for s in range(5, len(self.preference_data)):
      staff = self.schedule_data[s][0][2]
      for day in range(2, self.days_in_month+2):
        if self.preference_data[s][day] in [day_off1, day_off2]:
          objective += sum(shifts[(staff, day, n)] for n in range(num_shifts))
        elif self.preference_data[s][day] == night_shift:
          objective += shifts[(staff, day, 1)] + shifts[(staff, day, 2)]
        elif self.preference_data[s][day] == day_shift:
          objective += shifts[(staff, day, 0)] + shifts[(staff, day, 2)]
        elif self.preference_data[s][day] == evening_shift:
          objective += shifts[(staff, day, 0)] + shifts[(staff, day, 1)]

      if self.schedule_data[s][0][3] == shift_types[0]:
        objective += sum(shifts[(staff, day, 1)]
                         for day in range(2, self.days_in_month+2))
        objective += sum(shifts[(staff, day, 2)]
                         for day in range(2, self.days_in_month+2))

      if self.schedule_data[s][0][3] == shift_types[2]:
        objective += sum(shifts[(staff, day, 0)]
                         for day in range(2, self.days_in_month+2))
        objective += sum(shifts[(staff, day, 1)]
                         for day in range(2, self.days_in_month+2))
    model.Minimize(objective)

    # solve
    solver = cp_model.CpSolver()
    solver.parameters.linearization_level = 0
    solver.parameters.max_time_in_seconds = 20.0

    status = solver.Solve(model)
    self.set_optimize_status_signal.emit(solver.StatusName(status))

    #  print(solver.ObjectiveValue())

    try:
      for s in range(5, len(self.schedule_data)):
        staff = self.schedule_data[s][0][2]
        for day in range(2, self.days_in_month+2):
          if solver.Value(shifts[(staff, day, 0)]):
            self.schedule_data[s][day] = night_shift
          elif solver.Value(shifts[(staff, day, 1)]):
            self.schedule_data[s][day] = day_shift
          elif solver.Value(shifts[(staff, day, 2)]):
            self.schedule_data[s][day] = evening_shift
          else:
            # setting day off
            if self.preference_data[s][day] == business_travel:
              self.schedule_data[s][day] = business_travel
            elif self.preference_data[s][day] == day_off1:
              self.schedule_data[s][day] = day_off1
            elif self.preference_data[s][day] == day_off2:
              self.schedule_data[s][day] = day_off2
            else:
              self.schedule_data[s][day] = day_off1

      #  for s in range(5, len(self.preference_data)):
      #    staff = self.schedule_data[s][0][2]
      #    for day in range(2, self.days_in_month+2):
      #      if self.preference_data[s][day] in [day_off1, day_off2]:
      #        if solver.Value(shifts[(staff, day, 0)]) != 0 or \
      #            solver.Value(shifts[(staff, day, 1)]) != 0 or \
      #            solver.Value(shifts[(staff, day, 2)]) != 0:
      #          self.diff[s][day] = True
    except Exception:
      pass

    message = 'Statistics: conflicts: %d, branches: %d, wall time: %f' % (
      solver.NumConflicts(), solver.NumBranches(), solver.WallTime())
    self.status_message_signal.emit(message)

  def update_state(self):
    """
    compute the number of days off for each staff
    and fill in the empty cell as day shift for leader row
    """

    # update days off
    for s in range(5, len(self.schedule_data)):
      num_days_off = 0
      for day in range(2, self.days_in_month+2):
        if self.schedule_data[s][day] in unavailable:
          num_days_off += 1
      self.schedule_data[s][-1] = num_days_off

    # update leader
    for day in range(2, self.days_in_month+2):
      if self.schedule_data[4][day] == '':
        self.schedule_data[4][day] = day_shift

    self.highlight()

  def highlight(self):
    """
    create highlight if
      - green: the requested day off is not given
      - bold: the day has more than enough staff for the kind of shift
    """

    self.diff = [[False for _ in range(self.columnCount(0))]
                  for s in range(len(self.schedule_data))]

    for s in range(5, len(self.preference_data)):
      for day in range(2, self.days_in_month+2):
        if self.preference_data[s][day] in [day_off1, day_off2]:
          if self.schedule_data[s][day] not in [day_off1, day_off2]:
            self.diff[s][day] = True

    self.more = [[False for _ in range(self.columnCount(0))]
                 for s in range(len(self.schedule_data))]
    for day in range(2, self.days_in_month+2):
      required_night_shift = self.schedule_data[1][day]
      required_day_shift = self.schedule_data[2][day]
      required_evening_shift = self.schedule_data[3][day]

      night_shift_count = 0
      day_shift_count = 0
      evening_shift_count = 0
      for s in range(5, len(self.schedule_data)):
        if self.schedule_data[s][day] == night_shift:
          night_shift_count += 1
        elif self.schedule_data[s][day] == day_shift:
          day_shift_count += 1
        elif self.schedule_data[s][day] == evening_shift:
          evening_shift_count += 1

      if night_shift_count > required_night_shift:
        for s in range(5, len(self.schedule_data)):
          if self.schedule_data[s][day] == night_shift:
            self.more[s][day] = True

      if day_shift_count > required_day_shift:
        for s in range(5, len(self.schedule_data)):
          if self.schedule_data[s][day] == day_shift:
            self.more[s][day] = True

      if evening_shift_count > required_evening_shift:
        for s in range(5, len(self.schedule_data)):
          if self.schedule_data[s][day] == evening_shift:
            self.more[s][day] = True

  def export_json(self, filepath):
    obj = {
      'year': self.current_date.year,
      'month': self.current_date.month,
      'data': self.schedule_data,
    }

    try:
      with open(filepath, 'w') as f:
        json.dump(obj, f)
      return True
    except Exception:
      return False

  def export_csv(self, filepath):
    df = to_df(self.schedule_data, self.current_date,
               self.first_day, self.days_in_month)
    df.to_csv(filepath)

  def export_excel(self, filepath):
    df = to_df(self.schedule_data, self.current_date,
               self.first_day, self.days_in_month)
    df.to_excel(filepath)

  def import_json(self, filepath):
    with open(filepath) as f:
      obj = json.load(f)
    new_model_data = obj['data']

    # check if size is the same
    if len(new_model_data) != len(self.model_data):
      self.status_message_signal.emit(
        'fail to import data: incorrect number of rows')
      return False

    for i, row in enumerate(new_model_data):
      if len(row) != len(self.model_data[i]):
        self.status_message_signal.emit(
          'fail to import data: incorrect number of columns')
        return False

    for s in range(len(self.model_data)):
      for day in range(self.columnCount(0)):
        if not isinstance(self.model_data[s][day], type(new_model_data[s][day])):
          self.status_message_signal.emit(
            'fail to import data: not the same type')
          return False

    self.status_message_signal.emit('import success')
    self.beginResetModel()
    self.model_data = new_model_data
    self.endResetModel()
    self.save()
    return True

  def import_df(self, df):
    try:
      new_model_data = from_df(df)

      # check if size is the same
      if len(new_model_data) != len(self.schedule_data):
        return False

      for i, row in enumerate(new_model_data):
        if len(row) != len(self.schedule_data[i]):
          return False

      for s in range(len(self.model_data)):
        for day in range(self.columnCount(0)):
          if not isinstance(self.schedule_data[s][day], type(new_model_data[s][day])):
            return False

      self.status_message_signal.emit('import success')
      self.beginResetModel()
      self.schedule_data = new_model_data
      self.update_state()
      self.endResetModel()
      self.save()
      return True
    except Exception:
      self.status_message_signal.emit('fail to import data')

  def import_csv(self, filepath):
    df = pd.read_csv(filepath)
    self.import_df(df)

  def import_excel(self, filepath):
    df = pd.read_excel(filepath)
    self.import_df(df)

  def rowCount(self, parent):
    return len(self.schedule_data)

  def columnCount(self, parent):
    return self.days_in_month + 3

  def data(self, index, role):
    if role == Qt.DisplayRole:
      if index.column() == 0 and index.row() >= 4:
        return self.schedule_data[index.row()][index.column()][2]
      else:
        return self.schedule_data[index.row()][index.column()]
    elif role == Qt.EditRole:
      return self.schedule_data[index.row()][index.column()]
    elif role == Qt.BackgroundRole:
      leader_offset = 4
      staff_offset = 5
      if index.row() >= leader_offset:
        if index.row() >= staff_offset and index.column() == 0:
          if self.staffs[index.row()-staff_offset][-1] == shift_types[0]:
            return color1
          elif self.staffs[index.row()-staff_offset][-1] == shift_types[2]:
            return color2

        if self.schedule_data[index.row()][index.column()] == day_off1:
          return color3
        elif self.schedule_data[index.row()][index.column()] == day_off2:
          return color3
        elif self.schedule_data[index.row()][index.column()] == business_travel:
          return color3
        elif self.schedule_data[index.row()][index.column()] == night_shift:
          return color1
        elif self.schedule_data[index.row()][index.column()] == evening_shift:
          return color2
      elif index.row() == 1:
        return color1
      elif index.row() == 3:
        return color2
    elif role == Qt.ForegroundRole:
      if index.row() >= 5:
        if self.diff[index.row()][index.column()]:
          return color5
    elif role == Qt.FontRole:
      if index.row() >= 5:
        bold_font = QFont()
        bold_font.setBold(True)
        if self.more[index.row()][index.column()]:
          return bold_font
        if self.diff[index.row()][index.column()]:
          return bold_font
    elif role == Qt.TextAlignmentRole:
      return Qt.AlignCenter
    return None

  def setData(self, index, value, role):
    if role == Qt.EditRole:
      if isinstance(value, str):
        self.schedule_data[index.row()][index.column()] = value.upper()
      else:
        self.schedule_data[index.row()][index.column()] = value

      self.save()
      return True
    return False

  def headerData(self, col, orientation, role):
    if role == Qt.DisplayRole:
      if orientation == Qt.Horizontal:
        day_text = ['一', '二', '三', '四', '五', '六', '日']
        if col == 0:
          return 'Name'
        elif col == 1:
          return 'Limit'
        elif col > 1 and col <= self.days_in_month+1:
          text = day_text[(self.first_day + col - 2) % len(day_text)]
          return '%d/%d\n%s' % (self.current_date.month, col-1, text)
        else:
          return 'Total'
    return None

  def flags(self, index):
    if index.column() > 0:
      return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable
    return Qt.ItemIsEnabled | Qt.ItemIsSelectable


class MainWindow(QMainWindow, Ui_MainWindow):
  def __init__(self):
    QMainWindow.__init__(self)
    self.setupUi(self)

    # initialized values
    current_date = datetime.now()

    self.tab_widget.currentChanged.connect(self.handle_tab_change)

    # staffs tab
    self.staff_preference_combobox.addItems(shift_types)
    self.staff_preference_combobox.setCurrentIndex(1)
    self.staff_model = StaffModel(self)
    self.staff_view.setModel(self.staff_model)
    self.staff_model.status_message_signal.connect(self.show_status_message)
    self.staff_item_delegate = StaffItemDelegate(self)
    self.staff_view.setItemDelegate(self.staff_item_delegate)
    self.add_staff_button.clicked.connect(self.add_staff)
    self.delete_staff_button.clicked.connect(self.delete_staff)

    # leaders tab
    self.leader_model = LeaderModel(self)
    self.leader_view.setModel(self.leader_model)
    self.leader_model.status_message_signal.connect(self.show_status_message)
    self.add_leader_button.clicked.connect(self.add_leader)
    self.delete_leader_button.clicked.connect(self.delete_leader)

    # scheduling request tab
    self.request_month_combobox.addItems([str(m) for m in range(1, 13)])
    self.request_month_combobox.setCurrentIndex(current_date.month-1)
    self.request_year_lineedit.setText(str(current_date.year))
    self.request_model = RequestModel(self, current_date)
    self.request_view.setModel(self.request_model)
    self.request_model.status_message_signal.connect(
      self.show_status_message)
    self.load_request_button.clicked.connect(self.load_requests)
    self.request_night_shift_button.clicked.connect(
      lambda: self.request_shift(night_shift))
    self.request_day_shift_button.clicked.connect(
      lambda: self.request_shift(day_shift))
    self.request_evening_shift_button.clicked.connect(
      lambda: self.request_shift(evening_shift))
    self.request_day_off_button.clicked.connect(
      lambda: self.request_shift(day_off1))
    self.request_clear_button.clicked.connect(
      lambda: self.request_shift(''))
    self.export_request_button.clicked.connect(self.export_request)
    self.import_request_button.clicked.connect(self.import_request)

    self.request_view.setColumnWidth(0, 80)
    for col in range(1, self.request_model.columnCount(0)):
      self.request_view.setColumnWidth(col, 50)

    for row in range(0, self.request_model.rowCount(0)):
      self.request_view.setRowHeight(row, 30)

    # schedule tab
    self.schedule_month_combobox.addItems([str(m) for m in range(1, 13)])
    self.schedule_month_combobox.setCurrentIndex(current_date.month-1)
    self.schedule_year_lineedit.setText(str(current_date.year))
    self.schedule_model = ScheduleModel(self, current_date)
    self.schedule_view.setModel(self.schedule_model)
    self.schedule_model.status_message_signal.connect(
      self.show_status_message)
    self.schedule_model.set_optimize_status_signal.connect(
      self.set_optimize_status)
    self.load_schedule_button.clicked.connect(self.load_schedule)
    self.schedule_button.clicked.connect(self.schedule_model.optimize_asyn)
    self.work_day_constrain.textChanged.connect(self.schedule_model.set_work_day_constrain)
    self.day_off_constrain.textChanged.connect(self.schedule_model.set_day_off_contrain)
    self.export_schedule_button.clicked.connect(self.export_schedule)
    self.import_schedule_button.clicked.connect(self.import_schedule)

    self.schedule_view.setColumnWidth(0, 80)
    for col in range(1, self.schedule_model.columnCount(0)):
      self.schedule_view.setColumnWidth(col, 50)

    for row in range(0, self.schedule_model.rowCount(0)):
      self.schedule_view.setRowHeight(row, 30)

  def handle_tab_change(self, index):
    if index == 2:
      self.request_model.load_data()
    elif index == 3:
      self.schedule_model.load_data()

  def show_status_message(self, msg):
    self.statusbar.showMessage(msg, 20000)

  def show_error(self, message):
    message_box = QMessageBox(self)
    message_box.setIcon(QMessageBox.Critical)
    message_box.setText('Error')
    message_box.setInformativeText(message)
    message_box.setWindowTitle('Error')
    message_box.exec_()

  def add_staff(self):
    staff_id = self.staff_id_lineedit.text().strip()
    name = self.staff_name_lineedit.text().strip()
    preference_index = self.staff_preference_combobox.currentIndex()

    if len(name) > 0:
      if len(staff_id) > 0 and not staff_id.isnumeric():
        self.show_error('Invalid Staff ID')
      else:
        self.staff_model.add_staff(staff_id, name, preference_index)
        self.request_model.load_data()
        self.schedule_model.load_data()

        # reset inputs
        self.staff_id_lineedit.setText('')
        self.staff_name_lineedit.setText('')
        self.show_status_message('Entry Added')
    else:
      self.show_error('Need Staff Name')

  def confirmation(self, row_text, callback):
    message_box = QMessageBox(self)
    message_box.setIcon(QMessageBox.Question)
    message_box.setText('Delete Confirmation')
    message_box.setInformativeText(
      'Are you sure you want to delete these rows: ' + row_text)
    message_box.setWindowTitle('Delete Confirmation')
    message_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    #  message_box.buttonClicked.connect(delete_entry)
    ret = message_box.exec_()

    if ret == QMessageBox.Yes:
      callback()

  def get_selection(self, view):
    rows = []
    selection = []
    for index in view.selectionModel().selectedIndexes():
      if index.row() not in rows:
        rows.append(index.row())
        selection.append(index)

    rows = [str(r) for r in rows]
    return rows, selection

  def delete_staff(self):
    rows, selection = self.get_selection(self.staff_view)

    if len(selection) > 0:
      row_text = ','.join(rows)

      def callback():
        self.staff_model.delete_staff(selection)
        self.request_model.load_data()
        self.schedule_model.load_data()


      self.confirmation(row_text, callback)

  def add_leader(self):
    leader_id = self.leader_id_lineedit.text().strip()
    name = self.leader_name_lineedit.text().strip()

    if len(name) > 0:
      if len(leader_id) > 0 and not leader_id.isnumeric():
        self.show_error('Invalid Leader ID')
      else:
        self.leader_model.add_leader(leader_id, name)

        # reset inputs
        self.leader_id_lineedit.setText('')
        self.leader_name_lineedit.setText('')
        self.show_status_message('Entry Added')
    else:
      self.show_error('Name Required')

  def delete_leader(self):
    rows, selection = self.get_selection(self.leader_view)

    if len(selection) > 0:
      row_text = ','.join(rows)
      self.confirmation(row_text,
                        lambda: self.leader_model.delete_leader(selection))

  def load_requests(self):
    try:
      year = int(self.request_year_lineedit.text())
      month = self.request_month_combobox.currentIndex() + 1
      current_date = datetime(year, month, 1, 0, 0)

      self.request_model.set_current_date(current_date)
    except Exception as e:
      self.show_error(str(e))

  def load_schedule(self):
    try:
      year = int(self.schedule_year_lineedit.text())
      month = self.schedule_month_combobox.currentIndex() + 1
      current_date = datetime(year, month, 1, 0, 0)

      # need to make sure request is created
      # so let the request model load current date to create the data
      # and let schedule model load it
      c = self.request_model.current_date
      self.request_model.set_current_date(current_date)
      self.request_model.set_current_date(c)

      self.schedule_model.set_current_date(current_date)
    except Exception as e:
      self.show_error(str(e))

  def set_optimize_status(self, text):
    self.optimize_status.setText(text)
    if text == 'OPTIMAL':
      self.optimize_status.setStyleSheet('background-color: green; color: white;')
    elif text == 'FEASIBLE':
      self.optimize_status.setStyleSheet('background-color: yellow; color: black;')
    elif text == 'INFEASIBLE' or text == 'MODEL_INVALID' or text == 'UNKNOWN':
      self.optimize_status.setStyleSheet('background-color: red; color: white;')
    else:
      self.optimize_status.setStyleSheet('background-color: none; color: black;')

  def request_shift(self, shift):
    selection = self.request_view.selectionModel().selectedIndexes()
    if self.request_model.set_values(selection, shift):
      self.request_view.clearSelection()

  def export_request(self):
    result = QFileDialog.getSaveFileName(self,
      'Open Export File', '',
      'All Files (*.*);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;JSON Files (*.json)')

    if result[0]:
      filepath = result[0]
      if result[1] == 'All Files (*.*)' or \
          result[1] == 'Excel Files (*.xlsx *.xls)':
        if not filepath.lower().endswith('.xlsx') and \
            not filepath.lower().endswith('.xls'):
          filepath = filepath + '.xlsx'
        self.request_model.export_excel(filepath)
      elif result[1] == 'CSV Files (*.csv)':
        if not filepath.lower().endswith('.csv'):
          filepath = filepath + '.csv'
        self.request_model.export_csv(filepath)
      elif result[1] == 'JSON Files (*.json)':
        if not filepath.lower().endswith('.json'):
          filepath = filepath + '.json'
        self.request_model.export_json(filepath)

  def import_request(self):
    result = QFileDialog.getOpenFileName(self,
      'Open Export File', '', 'All Files (*.*)')

    if result[0]:
      filepath = result[0]
      if filepath.lower().endswith('.csv'):
        self.request_model.import_csv(filepath)
      elif filepath.lower().endswith('.xls') or \
          filepath.lower().endswith('.xlsx'):
        self.request_model.import_excel(filepath)
      elif filepath.lower().endswith('.json'):
        self.request_model.import_json(filepath)

  def export_schedule(self):
    result = QFileDialog.getSaveFileName(self,
      'Open Export File', '',
      'All Files (*.*);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;JSON Files (*.json)')

    if result[0]:
      filepath = result[0]
      if result[1] == 'All Files (*.*)' or \
          result[1] == 'Excel Files (*.xlsx *.xls)':
        if not filepath.lower().endswith('.xlsx') and \
            not filepath.lower().endswith('.xls'):
          filepath = filepath + '.xlsx'
        self.schedule_model.export_excel(filepath)
      elif result[1] == 'CSV Files (*.csv)':
        if not filepath.lower().endswith('.csv'):
          filepath = filepath + '.csv'
        self.schedule_model.export_csv(filepath)
      elif result[1] == 'JSON Files (*.json)':
        if not filepath.lower().endswith('.json'):
          filepath = filepath + '.json'
        self.schedule_model.export_json(filepath)

  def import_schedule(self):
    result = QFileDialog.getOpenFileName(self,
      'Open Export File', '', 'All Files (*.*)')

    if result[0]:
      filepath = result[0]
      if filepath.lower().endswith('.csv'):
        self.request_model.import_csv(filepath)
      elif filepath.lower().endswith('.xls') or \
          filepath.lower().endswith('.xlsx'):
        self.request_model.import_excel(filepath)
      elif filepath.lower().endswith('.json'):
        self.request_model.import_json(filepath)


def main():
  app = QApplication(sys.argv)
  app.aboutToQuit.connect(close_db)
  main_window = MainWindow()
  #  main_window.show()
  main_window.showMaximized()

  sys.exit(app.exec_())


if __name__ == '__main__':
  main()
