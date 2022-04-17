/* 0.0.2 наборщик номера по url схеме tel

wscript caller.min.js <action> [<number>]

<action>            - действие которое нужно выполнить
    reg             - зарегистрировать скрипт как обработчик
    unreg           - удалить регистрацию скрипта из обработчиков
    call            - выполнить набор номера
<number>            - телефонный номер или url для набора

*/

var caller = new App({
    login: "admin",         // имя пользователя для подключения к телефону
    password: "admin"       // пароль для подключения к телефону
});

// подключаем зависимые свойства приложения
(function (app, wsh, undefined) {// замыкаем чтобы не засорять глабальные объекты
    app.lib.extend(app, {// добавляем частный функционал приложения

        /**
         * Выполняет набор номера или изменяет регистрацию обработчика.
         * @param {string} action - Действие которое нужно выполнить.
         * @param {string} [number] - Телефонный номер или url для набора.
         * @returns {number} Номер ошибки или ноль.
         */

        action: function (action, number) {
            var value, data, shell, system, locator, service, response, item, items, ip, token,
                delim, flag, list, user, xhr, id = "call", link = "http://tl", error = 0;

            shell = new ActiveXObject("WScript.Shell");
            // корректируем номер для набора
            if (!error && number) {// если нужно выполнить
                value = app.lib.strim(number, ":", null, false, false);
                number = !value.indexOf("+7") ? "8" + value.substr(2) : value;
            };
            // формируем вспомогательные данные
            if (!error) {// если нету ошибок
                delim = "\\";// разделитель папок
                data = {// служебные данные
                    root: "HKEY_CLASSES_ROOT",
                    machine: "HKEY_LOCAL_MACHINE",
                    title: "Стационарный телефон",
                    original: "URL:Tel Protocol",
                    name: "URL Protocol",
                    protocol: "tel",
                    type: "REG_SZ",
                    id: "Caller"
                };
            };
            // выполняем регистрацию обработчика
            if (!error && "reg" == action) {// если нужно выполнить
                try {// пробуем выполнить действия
                    value = '"' + wsh.fullName + '" "' + wsh.scriptFullName + '" ' + id + ' "%1"';
                    shell.regWrite([data.root, data.id, ""].join(delim), data.title, data.type);
                    shell.regWrite([data.root, data.protocol, ""].join(delim), data.title, data.type);
                    shell.regWrite([data.root, data.protocol, data.name].join(delim), "", data.type);
                    shell.regWrite([data.root, data.id, "shell", "open", "command", ""].join(delim), value, data.type);
                    shell.regWrite([data.machine, "SOFTWARE", data.id, "Capabilities", "URLAssociations", data.protocol].join(delim), data.id, data.type);
                    shell.regWrite([data.machine, "SOFTWARE", "RegisteredApplications", data.id].join(delim), ["Software", data.id, "Capabilities"].join(delim), data.type);
                } catch (e) { error = 1; };
            };
            // удаляем регистрацию обработчика
            if (!error && "unreg" == action) {// если нужно выполнить
                try {// пробуем выполнить действия
                    shell.regWrite([data.root, data.id, ""].join(delim), data.original, data.type);
                    shell.regWrite([data.root, data.protocol, ""].join(delim), data.original, data.type);
                    shell.regWrite([data.root, data.protocol, data.name].join(delim), "", data.type);
                    try { shell.regDelete([data.root, data.id, "shell", "open", "command", ""].join(delim)); } catch (e) { };
                    try { shell.regDelete([data.root, data.id, "shell", "open", ""].join(delim)); } catch (e) { };
                    try { shell.regDelete([data.root, data.id, "shell", ""].join(delim)); } catch (e) { };
                    try { shell.regDelete([data.root, data.id, ""].join(delim)); } catch (e) { };
                    try { shell.regDelete([data.machine, "SOFTWARE", data.id, "Capabilities", "URLAssociations", ""].join(delim)); } catch (e) { };
                    try { shell.regDelete([data.machine, "SOFTWARE", data.id, "Capabilities", ""].join(delim)); } catch (e) { };
                    try { shell.regDelete([data.machine, "SOFTWARE", data.id, ""].join(delim)); } catch (e) { };
                    try { shell.regDelete([data.machine, "SOFTWARE", "RegisteredApplications", data.id].join(delim)); } catch (e) { };
                } catch (e) { error = 2; };
            };
            // выполняем телефонный звонок
            if (id == action && number) {// если нужно выполнить звонок
                // получаем wmi объект для взаимодействия
                if (!error) {// если нету ошибок
                    locator = new ActiveXObject("wbemScripting.Swbemlocator");
                    locator.security_.impersonationLevel = 3;// impersonate
                    try {// пробуем подключиться к нужному элименту
                        service = locator.connectServer(".", "root\\CIMV2");
                    } catch (e) { error = 3; };
                };
                // получаем ip адрес компьютера
                if (!error) {// если нету ошибок
                    response = service.execQuery(
                        "SELECT *" +
                        " FROM Win32_NetworkAdapterConfiguration" +
                        " WHERE ipEnabled = TRUE"
                    );
                    items = new Enumerator(response);
                    while (!items.atEnd()) {// пока не достигнут конец
                        item = items.item();// получаем очередной элимент коллекции
                        items.moveNext();// переходим к следующему элименту
                        // основной адрес 
                        if (null != item.ipAddress) {// если есть список ip адресов
                            list = item.ipAddress.toArray();// получаем очередной список
                            for (var i = 0, iLen = list.length; i < iLen && !ip; i++) {
                                value = list[i];// получаем очередное значение
                                if (value && ~value.indexOf('.')) ip = value;
                            };
                        };
                        // останавливаемся на первом элименте
                        break;
                    };
                    if (ip) {// если есть ip адрес компьютера
                    } else error = 4;
                };
                // определяем внутренний номер пользователя
                if (!error) {// если нету ошибок
                    system = new ActiveXObject("ADSystemInfo");
                    try {// пробуем подключиться к нужному элименту
                        user = GetObject("LDAP://" + system.userName);
                        value = user.get("telephoneNumber");
                    } catch (e) { error = 5; };
                };
                // вычисляем адрес для подключения к телефону
                if (!error) {// если нету ошибок
                    if (value) {// если есть внутренний номер
                        link += value;
                    } else error = 6;
                };
                // проходим авторизацию на телефоне
                if (!error) {// если нету ошибок
                    data = {// данные для запроса
                        username: app.val.login,
                        pwd: app.val.password,
                        jumpto: "features-remotecontrl"
                    };
                    xhr = app.lib.sjax("post", link + "/servlet?p=login&q=login", false, data);
                    if (200 == xhr.status) {// если запрос выполнен успешно
                        value = xhr.responseText;
                    } else error = 7;
                };
                // обрабатываем список разрешённых ip адресов
                if (!error) {// если нету ошибок
                    delim = ",";// разделитель значений
                    flag = true;// нужно ли ip адрес компьютера добавить в разрещённые
                    token = app.lib.strim(value, 'g_token = "', '"', false, false);
                    value = app.lib.strim(value, 'name="AURILimitIP"', '>', false, false);
                    list = app.lib.strim(value, 'value="', '"', false, false).split(delim);
                    for (var i = 0, iLen = list.length; i < iLen && flag; i++) {
                        value = list[i];// получаем очередное значение
                        if (value == ip) flag = false;
                    };
                    if (token) {// если удалось получить токен
                    } else error = 8;
                };
                // добавляем ip адрес компьютера в разрещённые
                if (!error && flag) {// если нужно выполнить
                    list.push(ip);// добавляем ip адрес в список
                    data = {// данные для запроса
                        "AURILimitIP": list.join(delim),
                        "token": token
                    };
                    xhr = app.lib.sjax("post", link + "/servlet?p=features-remotecontrl&q=write", false, data);
                    if (200 == xhr.status) {// если запрос выполнен успешно
                    } else error = 9;
                };
                // выполняем звонок
                if (!error) {// если нету ошибок
                    xhr = app.lib.sjax("get", link + "/servlet?p=contacts-callinfo&q=call&acc=0&num=" + number, true);
                    if ("call success" == xhr.responseText) {// если запрос выполнен успешно
                    } else error = 10;
                };
            };
            // отображаем информацию
            list = [];// обнуляем список значений
            list[1] = "Неудалось зарегистрировать протокол для звонков.";
            list[2] = "Неудалось удалить регистрацию протокола для звонков.";
            if (value = list[error]) wsh.echo(value);
            // возвращаем результат
            return error;
        },
        init: function () {// функция инициализации приложения
            var value, list = [], index = 0, error = 0;

            // формируем список параметров для вызова действия
            if (!error) {// если нету ошибок
                for (var i = index, iLen = wsh.arguments.length; i < iLen; i++) {
                    value = wsh.arguments(i);// получаем очередное значение
                    list.push(value);// добавляем значение в список
                };
            };
            // выполняем поддерживаемое действие
            if (!error) error = app.action.apply(app, list);
            // завершаем сценарий кодом
            wsh.quit(error);
        }
    });
})(caller, WSH);

// инициализируем приложение
caller.init();