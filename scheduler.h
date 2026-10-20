#ifndef SCHEDULER_H
#define SCHEDULER_H

#include <QMainWindow>
#include <QAbstractTableModel>
#include <QThread>
#include <QDate>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <vector>
#include <string>
#include <functional>

namespace Ui {
class MainWindow;
}

class StaffModel : public QAbstractTableModel {
    Q_OBJECT
public:
    explicit StaffModel(QObject *parent = nullptr);
    int rowCount(const QModelIndex &parent = QModelIndex()) const override;
    int columnCount(const QModelIndex &parent = QModelIndex()) const override;
    QVariant data(const QModelIndex &index, int role = Qt::DisplayRole) const override;
    bool setData(const QModelIndex &index, const QVariant &value, int role = Qt::EditRole) override;
    QVariant headerData(int section, Qt::Orientation orientation, int role = Qt::DisplayRole) const override;
    Qt::ItemFlags flags(const QModelIndex &index) const override;

    void load_data();
    void add_staff(int staff_id, const QString &name, int preference_index);
    void delete_staff(const QModelIndexList &selection);
    void set_preference(const QModelIndex &index, int current_index);

private:
    std::vector<std::vector<QVariant>> staff_data;
};

class LeaderModel : public QAbstractTableModel {
    Q_OBJECT
public:
    explicit LeaderModel(QObject *parent = nullptr);
    int rowCount(const QModelIndex &parent = QModelIndex()) const override;
    int columnCount(const QModelIndex &parent = QModelIndex()) const override;
    QVariant data(const QModelIndex &index, int role = Qt::DisplayRole) const override;
    bool setData(const QModelIndex &index, const QVariant &value, int role = Qt::EditRole) override;
    QVariant headerData(int section, Qt::Orientation orientation, int role = Qt::DisplayRole) const override;
    Qt::ItemFlags flags(const QModelIndex &index) const override;

    void load_data();
    void add_leader(int leader_id, const QString &name);
    void delete_leader(const QModelIndexList &selection);

private:
    std::vector<std::vector<QVariant>> leader_data;
};

class RequestModel : public QAbstractTableModel {
    Q_OBJECT
public:
    explicit RequestModel(QObject *parent = nullptr, QDate current_date = QDate::currentDate());
    int rowCount(const QModelIndex &parent = QModelIndex()) const override;
    int columnCount(const QModelIndex &parent = QModelIndex()) const override;
    QVariant data(const QModelIndex &index, int role = Qt::DisplayRole) const override;
    bool setData(const QModelIndex &index, const QVariant &value, int role = Qt::EditRole) override;
    QVariant headerData(int section, Qt::Orientation orientation, int role = Qt::DisplayRole) const override;
    Qt::ItemFlags flags(const QModelIndex &index) const override;

    void set_current_date(QDate current_date);
    void load_data();
    void update_states();
    void update_shift_count(int day);
    void save();
    void set_values(const QModelIndexList &selection, const QVariant &value);

private:
    QDate current_date;
    int days_in_month;
    std::vector<std::vector<QVariant>> request_data;
    std::vector<QVariant> staffs;
    std::vector<QVariant> leaders;
    std::vector<int> shift_count;
};

class ScheduleModel : public QAbstractTableModel {
    Q_OBJECT
public:
    explicit ScheduleModel(QObject *parent = nullptr, QDate current_date = QDate::currentDate());
    int rowCount(const QModelIndex &parent = QModelIndex()) const override;
    int columnCount(const QModelIndex &parent = QModelIndex()) const override;
    QVariant data(const QModelIndex &index, int role = Qt::DisplayRole) const override;
    bool setData(const QModelIndex &index, const QVariant &value, int role = Qt::EditRole) override;
    QVariant headerData(int section, Qt::Orientation orientation, int role = Qt::DisplayRole) const override;
    Qt::ItemFlags flags(const QModelIndex &index) const override;

    void set_work_day_constrain(int constrain);
    void set_day_off_constrain(int constrain);
    void set_current_date(QDate current_date);
    void load_staffs();
    void load_request();
    void load_data();
    void save();
    void load_previous_month_data();
    void optimize();
    void update_state();

signals:
    void status_message_signal(QString);
    void set_optimize_status_signal(QString);

private:
    QDate current_date;
    int days_in_month;
    int work_day_constrain;
    int day_off_constrain;
    std::vector<std::vector<QVariant>> schedule_data;
    std::vector<std::vector<QVariant>> preference_data;
    std::vector<QVariant> staffs;
    std::vector<QVariant> leaders;
};

class AsyncTask : public QThread {
    Q_OBJECT
public:
    AsyncTask(QObject *parent, std::function<void()> func);
    void run() override;

private:
    std::function<void()> func;
};

class MainWindow : public QMainWindow {
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void handle_tab_change(int index);
    void show_status_message(const QString &msg);
    void show_error(const QString &message);
    void add_staff();
    void delete_staff();
    void add_leader();
    void delete_leader();
    void load_requests();
    void load_schedule();
    void set_optimize_status(const QString &text);
    void optimize_asyn();

private:
    Ui::MainWindow *ui;
    StaffModel *staff_model;
    LeaderModel *leader_model;
    RequestModel *request_model;
    ScheduleModel *schedule_model;
    QDate current_date;

    QModelIndexList get_selection(QAbstractItemView *view);
    void close_db();
};

#endif // SCHEDULER_H
