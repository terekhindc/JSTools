/*ТАБЛИЦА ДЛЯ ОБРАБОТКИ КОЛЛЕКЦИИ, ПЕРЕДАВАЕМОЙ ИЗ БАЗЫ MONGODB
Сначала заполнить данными с коллекции с помощью Excel.Data.Set (collection)
Для создания таблицы на страничке с заголовками в нужном количестве: Excel.Create ('Пример1', Пример2', Пример3')
*/

let Excel = {};
Excel.isExist = $('div').is('#dataExcel');

Excel.Data = {};
Excel.Data.GetCollections = [];

/**
 * Заполнение табличных данных
 * @param {object} collection - коллекция, возвращаемая из запроса к MongoDB
 */
Excel.Data.Set = function (collection) {

    for (let unit in collection) {
        let _col = [];
        let _fields = Object.keys(collection[unit]);
        let _values = Object.values(collection[unit]);
        let _columnCount = 0;

        for (let key in _fields) {
            let _field = {};

            if (!_fields[key].includes('_')) {
                _columnCount++;
                _field[_fields[key]] = _values[key];
                _col.push(_field);
            }
        }

        Excel.Data.GetCollections.push(_col);

        Excel.Data.Row.count++;
        Excel.Data.Column.count = _columnCount;
    }
}

Excel.Data.Row = {
    count : 0,
    GetValue : function (index) {
        let _fields = {};
        for (let i=0; i<Excel.Data.GetCollections[index].length; i++) {
            _fields[i] = {
                field: Object.keys(Excel.Data.GetCollections[index][i]),
                value: Object.values(Excel.Data.GetCollections[index][i])
            }
        }

        return _fields;
    }
}

Excel.Data.Column = {
    count : 0,
    GetValue : function (rawIndex, columnIndex) {
        return Excel.Data.Row.GetValue(rawIndex)[columnIndex];
    },
    GetWidth : function (columnIndex) {
        let _width = 0;
        for (let i=0; i<this.count; i++) {
            if (_width < Excel.Data.Row.GetValue(i)[columnIndex].value[0].length) _width = Excel.Data.Row.GetValue(i)[columnIndex].value[0].length;
        }
        return _width;
    }
}

Excel.Test = function () {
    for (let row=0; row<Excel.Data.Row.count; row++) {
        for (let col=0; col<Excel.Data.Column.count; col++) {
            let result = Excel.Data.Column.GetValue(row, col);
            console.log('Поле: ' + result.field);
            console.log('Значение: ' + result.value);
        }
    }
}

/**
 * Отрисовка таблицы на странице
 * @param {...string} ...titles - список заголовков таблицы
 */
Excel.Create = function (...titles) {

    if (Excel.isExist === true) {
        Excel.html.remove();
    }

    Excel.html = document.createElement('div');
    Excel.html.setAttribute('class', 'dataExcel');

    let _tableWidth = Excel.Data.Column.count * 20;

    if (_tableWidth > 100) {
        _tableWidth = 100;
    }

    Excel.html.style.width = _tableWidth+'%';

    let _row = document.createElement('div');
    _row.setAttribute('class', 'excelTitles');
    Excel.html.appendChild(_row);

    for (let i = 0; i< arguments.length; i++) {

        let _cell = document.createElement('div');
        _cell.style.margin = '5px';
        _row.appendChild(_cell);

        let text = document.createElement('p');
        _cell.append(text);
        text.innerText = arguments[i];
    }

    Excel.AutoWidth();
    Excel.WriteIn();

    document.body.appendChild(Excel.html);
};

// Заполнение таблицы данными по строкам
Excel.WriteIn = function () {
    for (let row=0; row<Excel.Data.Row.count; row++) {
        Excel.AddRow(row);
    }
}

// Добавление новых строк
Excel.AddRow = function (rowIndex) {
    let _row = document.createElement('div');
    _row.setAttribute('class', 'row');
    Excel.html.appendChild(_row);

    for (let i = 0; i< Excel.Data.Column.count; i++) {
        let _cell = document.createElement('div');
        _cell.setAttribute('class', 'cell');
        _row.appendChild(_cell);
        _cell.style.width=Excel.AutoWidth(i)+'%';

        let text = document.createElement('p');
        _cell.append(text);
        text.innerText = Excel.Data.Column.GetValue(rowIndex, i).value;
    }
}

// Автоматический расчёт ширины столбцов
Excel.AutoWidth = function (indexColumn) {

    let _sum = 0;
    for (let i = 0; i < Excel.Data.Column.count; i++) {
        _sum += Excel.Data.Column.GetWidth(i);
    }

    let _onePercent = _sum / 100;
    let _widthColumnInPercent = []

    for (let i = 0; i < Excel.Data.Column.count; i++) {
        _widthColumnInPercent.push(Excel.Data.Column.GetWidth(i) / _onePercent);
    }

    return _widthColumnInPercent[indexColumn];
}