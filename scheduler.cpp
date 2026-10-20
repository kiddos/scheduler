#include "scheduler.h"
#include "ui_scheduler.h"
#include <QMessageBox>
#include <QInputDialog>
#include <QJsonDocument>
#include <QJsonArray>
#include <QJsonObject>
#include <QDebug>
#include <ortools/sat/cp_model.h>

static std::vector<QString> shift_types = {QString::fromUtf8("大夜 (PH)"), QString::fromUtf8("白班 (1~4)"), QString::fromUtf8("小夜 (4N)")};
static QString day_shift = QString::fromUtf8("白班 (1~4)");
static QString off_shift = QString::fromUtf8("休");

StaffModel::StaffModel(QObject *parent) : QAbstractTableModel(parent) {
    load_data();
}

int StaffModel::rowCount(const QModelIndex &parent) const {
    return staff_data.size();
}

int StaffModel::columnCount(const QModelIndex &parent) const {
    return 3;
}

QVariant StaffModel::data(const QModelIndex &index, int role) const {
    if (!index.isValid()) return QVariant();
    if (role == Qt::DisplayRole || role == Qt::EditRole) {
        return staff_data[index.row()][index.column()];
    }
    return QVariant();
}

bool StaffModel::setData(const QModelIndex &index, const QVariant &value, int role) {
    if (role == Qt::EditRole) {
        staff_data[index.row()][index.column()] = value;
        int id = staff_data[index.row()][3].toInt();
        QSqlQuery query;
        if (index.column() == 0) {
            query.prepare("UPDATE staffs SET staffId = :val WHERE id = :id");
        } else if (index.column() == 1) {
            query.prepare("UPDATE staffs SET name = :val WHERE id = :id");
        }
        query.bindValue(":val", value);
        query.bindValue(":id", id);
        query.exec();
        emit dataChanged(index, index, {Qt::EditRole, Qt::DisplayRole});
        return true;
    }
    return false;
}

QVariant StaffModel::headerData(int section, Qt::Orientation orientation, int role) const {
    if (role == Qt::DisplayRole && orientation == Qt::Horizontal) {
        if (section == 0) return QString::fromUtf8("員工代號");
        if (section == 1) return QString::fromUtf8("姓名");
        if (section == 2) return QString::fromUtf8("偏好班別");
    }
    return QVariant();
}

Qt::ItemFlags StaffModel::flags(const QModelIndex &index) const {
    if (index.column() == 2) return Qt::ItemIsEnabled | Qt::ItemIsSelectable;
    return Qt::ItemIsEditable | Qt::ItemIsEnabled | Qt::ItemIsSelectable;
}

void StaffModel::load_data() {
    beginResetModel();
    staff_data.clear();
    QSqlQuery query("SELECT staffId, name, preference, id FROM staffs");
    while (query.next()) {
        staff_data.push_back({query.value(0), query.value(1), query.value(2), query.value(3)});
    }
    endResetModel();
}

void StaffModel::add_staff(int staff_id, const QString &name, int preference_index) {
    QSqlQuery query;
    query.prepare("INSERT INTO staffs (staffId, name, preference) VALUES (:id, :name, :pref)");
    query.bindValue(":id", staff_id);
    query.bindValue(":name", name);
    query.bindValue(":pref", shift_types[preference_index]);
    query.exec();
    load_data();
}

void StaffModel::delete_staff(const QModelIndexList &selection) {
    if (selection.empty()) return;
    int row = selection.first().row();
    int id = staff_data[row][3].toInt();
    QSqlQuery query;
    query.prepare("DELETE FROM staffs WHERE id = :id");
    query.bindValue(":id", id);
    query.exec();
    load_data();
}

void StaffModel::set_preference(const QModelIndex &index, int current_index) {
    staff_data[index.row()][index.column()] = shift_types[current_index];
    int id = staff_data[index.row()][3].toInt();
    QSqlQuery query;
    query.prepare("UPDATE staffs SET preference = :pref WHERE id = :id");
    query.bindValue(":pref", shift_types[current_index]);
    query.bindValue(":id", id);
    query.exec();
    emit dataChanged(index, index);
}

LeaderModel::LeaderModel(QObject *parent) : QAbstractTableModel(parent) {
    load_data();
}

int LeaderModel::rowCount(const QModelIndex &parent) const {
    return leader_data.size();
}

int LeaderModel::columnCount(const QModelIndex &parent) const {
    return 2;
}

QVariant LeaderModel::data(const QModelIndex &index, int role) const {
    if (!index.isValid()) return QVariant();
    if (role == Qt::DisplayRole || role == Qt::EditRole) {
        return leader_data[index.row()][index.column()];
    }
    return QVariant();
}

bool LeaderModel::setData(const QModelIndex &index, const QVariant &value, int role) {
    if (role == Qt::EditRole) {
        leader_data[index.row()][index.column()] = value;
        int id = leader_data[index.row()][2].toInt();
        QSqlQuery query;
        if (index.column() == 0) {
            query.prepare("UPDATE leaders SET leaderId = :val WHERE id = :id");
        } else if (index.column() == 1) {
            query.prepare("UPDATE leaders SET name = :val WHERE id = :id");
        }
        query.bindValue(":val", value);
        query.bindValue(":id", id);
        query.exec();
        emit dataChanged(index, index, {Qt::EditRole, Qt::DisplayRole});
        return true;
    }
    return false;
}

QVariant LeaderModel::headerData(int section, Qt::Orientation orientation, int role) const {
    if (role == Qt::DisplayRole && orientation == Qt::Horizontal) {
        if (section == 0) return QString::fromUtf8("員工代號");
        if (section == 1) return QString::fromUtf8("姓名");
    }
    return QVariant();
}

Qt::ItemFlags LeaderModel::flags(const QModelIndex &index) const {
    return Qt::ItemIsEditable | Qt::ItemIsEnabled | Qt::ItemIsSelectable;
}

void LeaderModel::load_data() {
    beginResetModel();
    leader_data.clear();
    QSqlQuery query("SELECT leaderId, name, id FROM leaders");
    while (query.next()) {
        leader_data.push_back({query.value(0), query.value(1), query.value(2)});
    }
    endResetModel();
}

void LeaderModel::add_leader(int leader_id, const QString &name) {
    QSqlQuery query;
    query.prepare("INSERT INTO leaders (leaderId, name) VALUES (:id, :name)");
    query.bindValue(":id", leader_id);
    query.bindValue(":name", name);
    query.exec();
    load_data();
}

void LeaderModel::delete_leader(const QModelIndexList &selection) {
    if (selection.empty()) return;
    int row = selection.first().row();
    int id = leader_data[row][2].toInt();
    QSqlQuery query;
    query.prepare("DELETE FROM leaders WHERE id = :id");
    query.bindValue(":id", id);
    query.exec();
    load_data();
}

RequestModel::RequestModel(QObject *parent, QDate current_date) : QAbstractTableModel(parent) {
    set_current_date(current_date);
}

void RequestModel::set_current_date(QDate date) {
    current_date = date;
    days_in_month = date.daysInMonth();
    load_data();
}

void RequestModel::load_data() {
    beginResetModel();
    request_data.clear();
    staffs.clear();
    leaders.clear();

    QSqlQuery q1("SELECT staffId, name, preference FROM staffs");
    while (q1.next()) {
        staffs.push_back(QVariantList{q1.value(0), q1.value(1), q1.value(2)});
    }

    QSqlQuery q2("SELECT leaderId, name FROM leaders");
    while (q2.next()) {
        leaders.push_back(QVariantList{q2.value(0), q2.value(1)});
    }

    QSqlQuery query;
    query.prepare("SELECT data FROM requests WHERE year = :y AND month = :m");
    query.bindValue(":y", current_date.year());
    query.bindValue(":m", current_date.month());
    query.exec();

    if (query.next()) {
        QJsonDocument doc = QJsonDocument::fromJson(query.value(0).toString().toUtf8());
        QJsonArray arr = doc.array();
        for (int i = 0; i < arr.size(); ++i) {
            std::vector<QVariant> row;
            QJsonArray rowArr = arr[i].toArray();
            for (int j = 0; j < rowArr.size(); ++j) {
                if (j == 0) {
                    if (rowArr[j].isArray()) {
                        QJsonArray nested = rowArr[j].toArray();
                        QVariantList lst;
                        for(int k=0; k<nested.size(); ++k) lst.append(nested[k].toVariant());
                        row.push_back(lst);
                    } else {
                        row.push_back(rowArr[j].toVariant());
                    }
                } else {
                    row.push_back(rowArr[j].toVariant());
                }
            }
            request_data.push_back(row);
        }
    } else {
        request_data.push_back({QString::fromUtf8("上班護理師人數")});
        request_data.back().resize(days_in_month + 1, 0);

        request_data.push_back({QString::fromUtf8("大夜班 (PH)")});
        request_data.back().resize(days_in_month + 1, 0);

        request_data.push_back({QString::fromUtf8("白班 (1~4)")});
        request_data.back().resize(days_in_month + 1, 0);

        request_data.push_back({QString::fromUtf8("小夜班 (4N)")});
        request_data.back().resize(days_in_month + 1, 0);

        for (const auto &leader : leaders) {
            std::vector<QVariant> row = {leader};
            row.resize(days_in_month + 1, "");
            request_data.push_back(row);
        }

        for (const auto &staff : staffs) {
            std::vector<QVariant> row = {staff};
            row.resize(days_in_month + 1, "");
            request_data.push_back(row);
        }
    }
    endResetModel();
}

int RequestModel::rowCount(const QModelIndex &parent) const {
    return request_data.size();
}

int RequestModel::columnCount(const QModelIndex &parent) const {
    return days_in_month + 1;
}

QVariant RequestModel::data(const QModelIndex &index, int role) const {
    if (!index.isValid()) return QVariant();
    int row = index.row();
    int col = index.column();

    if (role == Qt::DisplayRole || role == Qt::EditRole) {
        if (col == 0 && row >= 4) {
            QVariantList lst = request_data[row][0].toList();
            return QString("%1 - %2").arg(lst[0].toString()).arg(lst[1].toString());
        }
        return request_data[row][col];
    } else if (role == Qt::BackgroundRole) {
        if (col == 0) return QBrush(QColor(220, 220, 220));
        int day_of_week = QDate(current_date.year(), current_date.month(), col).dayOfWeek();
        if (day_of_week == 6 || day_of_week == 7) {
            return QBrush(QColor(255, 204, 204));
        }
    }
    return QVariant();
}

bool RequestModel::setData(const QModelIndex &index, const QVariant &value, int role) {
    if (role == Qt::EditRole) {
        request_data[index.row()][index.column()] = value;
        emit dataChanged(index, index, {Qt::EditRole, Qt::DisplayRole});
        save();
        return true;
    }
    return false;
}

QVariant RequestModel::headerData(int section, Qt::Orientation orientation, int role) const {
    if (role == Qt::DisplayRole) {
        if (orientation == Qt::Horizontal) {
            if (section == 0) return QString::fromUtf8("員工 / 日期");
            return QString::number(section);
        }
    }
    return QVariant();
}

Qt::ItemFlags RequestModel::flags(const QModelIndex &index) const {
    if (index.column() == 0) return Qt::ItemIsEnabled;
    return Qt::ItemIsEditable | Qt::ItemIsEnabled | Qt::ItemIsSelectable;
}

void RequestModel::save() {
    QJsonArray arr;
    for (const auto &row : request_data) {
        QJsonArray rowArr;
        for (int i=0; i<row.size(); ++i) {
            if (i == 0 && row[0].type() == QVariant::List) {
                rowArr.append(QJsonArray::fromVariantList(row[0].toList()));
            } else {
                rowArr.append(QJsonValue::fromVariant(row[i]));
            }
        }
        arr.append(rowArr);
    }
    QString dataStr = QJsonDocument(arr).toJson(QJsonDocument::Compact);

    QSqlQuery checkQuery;
    checkQuery.prepare("SELECT id FROM requests WHERE year = :y AND month = :m");
    checkQuery.bindValue(":y", current_date.year());
    checkQuery.bindValue(":m", current_date.month());
    checkQuery.exec();

    QSqlQuery query;
    if (checkQuery.next()) {
        query.prepare("UPDATE requests SET data = :d WHERE year = :y AND month = :m");
    } else {
        query.prepare("INSERT INTO requests (year, month, data) VALUES (:y, :m, :d)");
    }
    query.bindValue(":y", current_date.year());
    query.bindValue(":m", current_date.month());
    query.bindValue(":d", dataStr);
    query.exec();
}

void RequestModel::update_states() {}
void RequestModel::update_shift_count(int day) {}
void RequestModel::set_values(const QModelIndexList &selection, const QVariant &value) {
    for(auto &idx : selection) {
        setData(idx, value);
    }
}

ScheduleModel::ScheduleModel(QObject *parent, QDate current_date) : QAbstractTableModel(parent) {
    this->work_day_constrain = 7;
    this->day_off_constrain = 1;
    set_current_date(current_date);
}

void ScheduleModel::set_work_day_constrain(int constrain) { work_day_constrain = constrain; }
void ScheduleModel::set_day_off_constrain(int constrain) { day_off_constrain = constrain; }

void ScheduleModel::set_current_date(QDate date) {
    current_date = date;
    days_in_month = date.daysInMonth();
    load_data();
}

void ScheduleModel::load_staffs() {
    staffs.clear();
    leaders.clear();
    QSqlQuery q1("SELECT staffId, name, preference FROM staffs");
    while (q1.next()) staffs.push_back(QVariantList{q1.value(0), q1.value(1), q1.value(2)});
    QSqlQuery q2("SELECT leaderId, name FROM leaders");
    while (q2.next()) leaders.push_back(QVariantList{q2.value(0), q2.value(1)});
}

void ScheduleModel::load_request() {
    preference_data.clear();
    QSqlQuery query;
    query.prepare("SELECT data FROM requests WHERE year = :y AND month = :m");
    query.bindValue(":y", current_date.year());
    query.bindValue(":m", current_date.month());
    query.exec();
    if (query.next()) {
        QJsonArray arr = QJsonDocument::fromJson(query.value(0).toString().toUtf8()).array();
        for (int i = 0; i < arr.size(); ++i) {
            std::vector<QVariant> row;
            QJsonArray rowArr = arr[i].toArray();
            for (int j = 0; j < rowArr.size(); ++j) {
                if (j == 0 && rowArr[j].isArray()) {
                    QJsonArray nested = rowArr[j].toArray();
                    QVariantList lst;
                    for(int k=0; k<nested.size(); ++k) lst.append(nested[k].toVariant());
                    row.push_back(lst);
                } else {
                    row.push_back(rowArr[j].toVariant());
                }
            }
            preference_data.push_back(row);
        }
    }
}

void ScheduleModel::load_data() {
    beginResetModel();
    load_staffs();
    load_request();

    QSqlQuery query;
    query.prepare("SELECT data FROM schedules WHERE year = :y AND month = :m");
    query.bindValue(":y", current_date.year());
    query.bindValue(":m", current_date.month());
    query.exec();

    if (query.next()) {
        QJsonArray arr = QJsonDocument::fromJson(query.value(0).toString().toUtf8()).array();
        schedule_data.clear();
        for (int i = 0; i < arr.size(); ++i) {
            std::vector<QVariant> row;
            QJsonArray rowArr = arr[i].toArray();
            for (int j = 0; j < rowArr.size(); ++j) {
                if (j == 0 && rowArr[j].isArray()) {
                    QJsonArray nested = rowArr[j].toArray();
                    QVariantList lst;
                    for(int k=0; k<nested.size(); ++k) lst.append(nested[k].toVariant());
                    row.push_back(lst);
                } else {
                    row.push_back(rowArr[j].toVariant());
                }
            }
            schedule_data.push_back(row);
        }
    } else {
        schedule_data = preference_data;
    }
    endResetModel();
}

int ScheduleModel::rowCount(const QModelIndex &parent) const {
    return schedule_data.size();
}

int ScheduleModel::columnCount(const QModelIndex &parent) const {
    return days_in_month + 1;
}

QVariant ScheduleModel::data(const QModelIndex &index, int role) const {
    if (!index.isValid()) return QVariant();
    int row = index.row();
    int col = index.column();

    if (role == Qt::DisplayRole || role == Qt::EditRole) {
        if (col == 0 && row >= 4) {
            QVariantList lst = schedule_data[row][0].toList();
            return QString("%1 - %2").arg(lst[0].toString()).arg(lst[1].toString());
        }
        return schedule_data[row][col];
    }
    return QVariant();
}

bool ScheduleModel::setData(const QModelIndex &index, const QVariant &value, int role) {
    if (role == Qt::EditRole) {
        schedule_data[index.row()][index.column()] = value;
        emit dataChanged(index, index, {Qt::EditRole, Qt::DisplayRole});
        save();
        return true;
    }
    return false;
}

QVariant ScheduleModel::headerData(int section, Qt::Orientation orientation, int role) const {
    if (role == Qt::DisplayRole && orientation == Qt::Horizontal) {
        if (section == 0) return QString::fromUtf8("員工 / 日期");
        return QString::number(section);
    }
    return QVariant();
}

Qt::ItemFlags ScheduleModel::flags(const QModelIndex &index) const {
    if (index.column() == 0 || index.row() < 4) return Qt::ItemIsEnabled;
    return Qt::ItemIsEditable | Qt::ItemIsEnabled | Qt::ItemIsSelectable;
}

void ScheduleModel::save() {
    QJsonArray arr;
    for (const auto &row : schedule_data) {
        QJsonArray rowArr;
        for (int i=0; i<row.size(); ++i) {
            if (i == 0 && row[0].type() == QVariant::List) {
                rowArr.append(QJsonArray::fromVariantList(row[0].toList()));
            } else {
                rowArr.append(QJsonValue::fromVariant(row[i]));
            }
        }
        arr.append(rowArr);
    }
    QString dataStr = QJsonDocument(arr).toJson(QJsonDocument::Compact);

    QSqlQuery checkQuery;
    checkQuery.prepare("SELECT id FROM schedules WHERE year = :y AND month = :m");
    checkQuery.bindValue(":y", current_date.year());
    checkQuery.bindValue(":m", current_date.month());
    checkQuery.exec();

    QSqlQuery query;
    if (checkQuery.next()) {
        query.prepare("UPDATE schedules SET data = :d WHERE year = :y AND month = :m");
    } else {
        query.prepare("INSERT INTO schedules (year, month, data) VALUES (:y, :m, :d)");
    }
    query.bindValue(":y", current_date.year());
    query.bindValue(":m", current_date.month());
    query.bindValue(":d", dataStr);
    query.exec();
}

void ScheduleModel::load_previous_month_data() {}
void ScheduleModel::update_state() {}

void ScheduleModel::optimize() {
    using namespace operations_research;
    using namespace operations_research::sat;

    emit set_optimize_status_signal("Optimization in progress...");

    CpModelBuilder cp_model;

    int num_shifts = 3;
    std::vector<QString> staff_names;
    std::vector<QString> staff_prefs;
    for(size_t s=4 + leaders.size(); s<schedule_data.size(); ++s) {
        QVariantList lst = schedule_data[s][0].toList();
        staff_names.push_back(lst[1].toString());
        staff_prefs.push_back(lst[2].toString());
    }
    int num_staffs = staff_names.size();

    std::map<std::tuple<int, int, int>, BoolVar> shifts;
    for(int s=0; s<num_staffs; ++s) {
        for(int d=1; d<=days_in_month; ++d) {
            for(int n=0; n<num_shifts; ++n) {
                shifts[{s, d, n}] = cp_model.NewBoolVar().WithName(QString("shift_%1_%2_%3").arg(s).arg(d).arg(n).toStdString());
            }
        }
    }

    // Required staff per day
    for(int d=1; d<=days_in_month; ++d) {
        for(int n=0; n<num_shifts; ++n) {
            std::vector<BoolVar> vars;
            for(int s=0; s<num_staffs; ++s) vars.push_back(shifts[{s, d, n}]);
            int required = schedule_data[n+1][d].toInt();
            cp_model.AddGreaterOrEqual(LinearExpr::Sum(vars), required);
        }
    }

    // Constraints per staff
    for(int s=0; s<num_staffs; ++s) {
        for(int d=1; d<=days_in_month; ++d) {
            std::vector<BoolVar> day_shifts;
            for(int n=0; n<num_shifts; ++n) day_shifts.push_back(shifts[{s, d, n}]);
            cp_model.AddLessOrEqual(LinearExpr::Sum(day_shifts), 1);

            if (d < days_in_month) {
                cp_model.AddLessOrEqual(LinearExpr::Sum({shifts[{s, d, 1}], shifts[{s, d, 2}], shifts[{s, d+1, 0}]}), 1);
                cp_model.AddLessOrEqual(LinearExpr::Sum({shifts[{s, d, 2}], shifts[{s, d+1, 0}], shifts[{s, d+1, 1}]}), 1);
            }
        }

        // Month total
        std::vector<BoolVar> month_shifts;
        for(int d=1; d<=days_in_month; ++d) {
            for(int n=0; n<num_shifts; ++n) month_shifts.push_back(shifts[{s, d, n}]);
        }
        cp_model.AddLessOrEqual(LinearExpr::Sum(month_shifts), days_in_month);

        // Preference
        int pref_count = 16;
        if (staff_prefs[s] == shift_types[0]) {
            std::vector<BoolVar> pref_shifts;
            for(int d=1; d<=days_in_month; ++d) pref_shifts.push_back(shifts[{s, d, 0}]);
            cp_model.AddGreaterOrEqual(LinearExpr::Sum(pref_shifts), pref_count);
        } else if (staff_prefs[s] == shift_types[2]) {
            std::vector<BoolVar> pref_shifts;
            for(int d=1; d<=days_in_month; ++d) pref_shifts.push_back(shifts[{s, d, 2}]);
            cp_model.AddGreaterOrEqual(LinearExpr::Sum(pref_shifts), pref_count);
        }

        // Work day constraints
        for(int d=1; d<=days_in_month - work_day_constrain; ++d) {
            std::vector<BoolVar> work_days;
            for(int w=0; w<=work_day_constrain; ++w) {
                for(int n=0; n<num_shifts; ++n) work_days.push_back(shifts[{s, d+w, n}]);
            }
            cp_model.AddLessOrEqual(LinearExpr::Sum(work_days), work_day_constrain);
        }
    }

    // Pre-assigned shifts (requests)
    for(size_t i=4 + leaders.size(); i<schedule_data.size(); ++i) {
        int s = i - 4 - leaders.size();
        for(int d=1; d<=days_in_month; ++d) {
            QString req = preference_data[i][d].toString();
            if (req == off_shift) {
                for(int n=0; n<num_shifts; ++n) cp_model.AddEquality(shifts[{s, d, n}], 0);
            } else if (req == shift_types[0]) {
                cp_model.AddEquality(shifts[{s, d, 0}], 1);
            } else if (req == shift_types[1]) {
                cp_model.AddEquality(shifts[{s, d, 1}], 1);
            } else if (req == shift_types[2]) {
                cp_model.AddEquality(shifts[{s, d, 2}], 1);
            }
        }
    }

    const CpSolverResponse response = Solve(cp_model.Build());

    if (response.status() == CpSolverStatus::OPTIMAL || response.status() == CpSolverStatus::FEASIBLE) {
        beginResetModel();
        for(int s=0; s<num_staffs; ++s) {
            for(int d=1; d<=days_in_month; ++d) {
                QString req = preference_data[s + 4 + leaders.size()][d].toString();
                if (req == off_shift || req == shift_types[0] || req == shift_types[1] || req == shift_types[2]) {
                    schedule_data[s + 4 + leaders.size()][d] = req;
                } else {
                    bool worked = false;
                    for(int n=0; n<num_shifts; ++n) {
                        if (SolutionIntegerValue(response, shifts[{s, d, n}]) == 1) {
                            schedule_data[s + 4 + leaders.size()][d] = shift_types[n];
                            worked = true;
                            break;
                        }
                    }
                    if (!worked) schedule_data[s + 4 + leaders.size()][d] = off_shift;
                }
            }
        }
        endResetModel();
        save();
        emit set_optimize_status_signal("Optimization completed!");
    } else {
        emit set_optimize_status_signal("Optimization failed (no solution).");
    }
}

AsyncTask::AsyncTask(QObject *parent, std::function<void()> func) : QThread(parent), func(func) {}
void AsyncTask::run() { func(); }

MainWindow::MainWindow(QWidget *parent) : QMainWindow(parent), ui(new Ui::MainWindow) {
    ui->setupUi(this);

    QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE");
    db.setDatabaseName("db.sqlite3");
    if (!db.open()) {
        QMessageBox::critical(this, "Error", "Cannot open database");
    }

    QSqlQuery query;
    query.exec("CREATE TABLE IF NOT EXISTS staffs (id INTEGER PRIMARY KEY AUTOINCREMENT, staffId INTEGER, name TEXT NOT NULL, preference TEXT NOT NULL)");
    query.exec("CREATE TABLE IF NOT EXISTS leaders (id INTEGER PRIMARY KEY AUTOINCREMENT, leaderId INTEGER, name TEXT NOT NULL)");
    query.exec("CREATE TABLE IF NOT EXISTS requests (id INTEGER PRIMARY KEY AUTOINCREMENT, year INTEGER NOT NULL, month INTEGER NOT NULL, data TEXT NOT NULL)");
    query.exec("CREATE TABLE IF NOT EXISTS schedules (id INTEGER PRIMARY KEY AUTOINCREMENT, year INTEGER NOT NULL, month INTEGER NOT NULL, data TEXT NOT NULL)");

    current_date = QDate::currentDate().addMonths(1);

    staff_model = new StaffModel(this);
    ui->staff_table->setModel(staff_model);

    leader_model = new LeaderModel(this);
    ui->leader_table->setModel(leader_model);

    request_model = new RequestModel(this, current_date);
    ui->request_table->setModel(request_model);

    schedule_model = new ScheduleModel(this, current_date);
    ui->schedule_table->setModel(schedule_model);

    connect(ui->tab_widget, &QTabWidget::currentChanged, this, &MainWindow::handle_tab_change);
    connect(ui->add_staff_btn, &QPushButton::clicked, this, &MainWindow::add_staff);
    connect(ui->del_staff_btn, &QPushButton::clicked, this, &MainWindow::delete_staff);
    connect(ui->add_leader_btn, &QPushButton::clicked, this, &MainWindow::add_leader);
    connect(ui->del_leader_btn, &QPushButton::clicked, this, &MainWindow::delete_leader);

    connect(ui->run_btn, &QPushButton::clicked, this, &MainWindow::optimize_asyn);
    connect(schedule_model, &ScheduleModel::set_optimize_status_signal, this, &MainWindow::set_optimize_status);
}

MainWindow::~MainWindow() {
    delete ui;
    close_db();
}

void MainWindow::close_db() {
    QSqlDatabase::database().close();
}

void MainWindow::handle_tab_change(int index) {
    if (index == 2) load_requests();
    else if (index == 3) load_schedule();
}

void MainWindow::show_status_message(const QString &msg) {
    ui->statusbar->showMessage(msg, 5000);
}

void MainWindow::show_error(const QString &message) {}

QModelIndexList MainWindow::get_selection(QAbstractItemView *view) {
    return view->selectionModel()->selectedIndexes();
}

void MainWindow::add_staff() {
    staff_model->add_staff(0, QString::fromUtf8("新員工"), 1);
}

void MainWindow::delete_staff() {
    staff_model->delete_staff(get_selection(ui->staff_table));
}

void MainWindow::add_leader() {
    leader_model->add_leader(0, QString::fromUtf8("新主管"));
}

void MainWindow::delete_leader() {
    leader_model->delete_leader(get_selection(ui->leader_table));
}

void MainWindow::load_requests() {
    int m = ui->request_month_combobox->currentIndex() + 1;
    current_date = QDate(current_date.year(), m, 1);
    request_model->set_current_date(current_date);
}

void MainWindow::load_schedule() {
    int m = ui->schedule_month_combobox->currentIndex() + 1;
    current_date = QDate(current_date.year(), m, 1);
    schedule_model->set_current_date(current_date);
}

void MainWindow::set_optimize_status(const QString &text) {
    ui->optimize_status_label->setText(text);
}

void MainWindow::optimize_asyn() {
    schedule_model->set_work_day_constrain(ui->work_day_constrain_combobox->currentText().toInt());
    schedule_model->set_day_off_constrain(ui->day_off_constrain_combobox->currentText().toInt());
    AsyncTask *task = new AsyncTask(this, [this](){ schedule_model->optimize(); });
    task->start();
}
