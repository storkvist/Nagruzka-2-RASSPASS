<?php

//var_dump(substr('ДЦим-2-1', 1));
//die();

$RASSPASS = 'C:\\2012_2013 1\\Raspis1.mdb';
$NAGRUZKA = 'VITALY-XP\SQLEXPRESS';

include 'adodb517/adodb.inc.php';

/**
 * @var ADOConnection $nagruzka
 */
$nagruzka = ADONewConnection('odbc_mssql');
$nagruzka->debug = true;
$nagruzka->SetFetchMode(ADODB_FETCH_ASSOC);

$dsn = "DRIVER={SQL Server};SERVER={$NAGRUZKA};DATABASE=Деканат;";
$nagruzka->Connect($dsn);

/**
 * @var ADOConnection $rasspass
 */
$rasspass = ADONewConnection('access');
$rasspass->debug = true;
$rasspass->SetFetchMode(ADODB_FETCH_ASSOC);


$dsn = "Driver={Microsoft Access Driver (*.mdb)};Dbq={$RASSPASS};Uid=Admin;";
$rasspass->Connect($dsn);

$query = <<<EOT
SELECT Дисциплина, Группа, Студентов, Специальность
FROM dbo.Нагрузка
WHERE
    КодТипаПлана IN (1, 2, 3)
    AND ВидЗанятий IN ('Лек', 'Лаб', 'Пр')
    AND Студентов > 0
    AND Дисциплина NOT LIKE '%осэкзамен%'
    AND (
        (Курс = 1 AND Семестр = 1)
        OR
        (Курс = 2 AND Семестр = 3)
        OR
        (Курс = 3 AND Семестр = 5)
        OR
        (Курс = 4 AND Семестр = 7)
        OR
        (Курс = 5 AND Семестр = 9)
        OR
        (Курс = 6 AND Семестр = 11)
    )
EOT;

$result = $nagruzka->Execute($query);

$flows = array(
    'Д' => array(),
    'В' => array(),
    'З' => array()
);
$specs = array();
foreach ($result->GetRows() as $row) {
    $form = $row['Группа'][0];

    $groupInfo = explode('-', substr($row['Группа'], 1));
    $flow = $groupInfo[0];

    if (!in_array($row['Специальность'], $specs)) {
        $specs[] = $row['Специальность'];
    }

    if (isset($flows[$form][$flow])) {
        $flows[$form][$flow][] = $row;
    } else {
        $flows[$form][$flow] = array($row);
    }
}

/*asort($specs);
foreach ($specs as $spec) {
    echo $spec . '<br>';
}
die();*/

foreach ($flows as $form) {
    foreach ($form as $flow => $groups) {
        var_dump($flow); echo '<br>';
    }
    echo "===============<br>";
}

var_dump(count($flows['Д']));
var_dump(count($flows['В']));
var_dump(count($flows['З']));
die();

$subjects = array();
foreach ($result->GetRows() as $row) {
    $subjects[] = $row['Дисциплина'];
}

/**
 * Убираем всякую хрень из названий предметов.
 */
foreach ($subjects as $index => $name) {
    $subjects[$index] = str_replace(
        array(
            ', п/г 1', ', п/г 2', ', п/г 3', ', п/г 4', ', п/г 5', ', п/г 6',
            ', п/г 7', ', п/г 8', ', преп 1', ', преп 3', ', преп 2',
            ', преп 4', ', преп 5', ', преп 6', ', преп 7', ', преп 8',
            ', преп 9', ', преп 10', ', Английский', ', Немецкий'
        ),
        '', $name
    );
}
$subjects = array_unique($subjects, SORT_STRING);

/**
 * Переносим в RASSPASS список дисциплин.
 */
$rasspass->Execute('DELETE FROM Дисциплины');
foreach ($subjects as $subject) {
    $rasspass->Execute(
        "INSERT INTO Дисциплины (Дисциплина) VALUES ('{$subject}')"
    );
}

$result = $nagruzka->Execute(
    "SELECT Группа, Студентов FROM dbo.Нагрузка WHERE ВидЗанятий IN ('Лек', 'Лаб', 'Пр') AND Дисциплина NOT LIKE '%осэкзамен%' AND Студентов > 0"
);
$groups = array();
foreach ($result->GetRows() as $row) {
    if (isset($groups[$row['Группа']])) {
        if ($groups[$row['Группа']] != $row['Студентов']) {
            var_dump($row['Группа']);
            die();
        }
    } else {
        $groups[$row['Группа']] = $row['Студентов'];
    }
}

$specs = array();
$flows = array();

var_dump($groups);
var_dump(count($groups));
