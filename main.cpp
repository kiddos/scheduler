#include <QApplication>
#include "scheduler.h"

int main(int argc, char *argv[]) {
    QApplication app(argc, argv);
    MainWindow main_window;
    main_window.showMaximized();
    return app.exec();
}
