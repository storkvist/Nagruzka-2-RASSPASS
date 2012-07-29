<?php

//var_dump(substr('����-2-1', 1));
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

$dsn = "DRIVER={SQL Server};SERVER={$NAGRUZKA};DATABASE=�������;";
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
SELECT ����������, ������, ���������, �������������
FROM dbo.��������
WHERE
    ������������ IN (1, 2, 3)
    AND ���������� IN ('���', '���', '��')
    AND ��������� > 0
    AND ���������� NOT LIKE '%���������%'
    AND (
        (���� = 1 AND ������� = 1)
        OR
        (���� = 2 AND ������� = 3)
        OR
        (���� = 3 AND ������� = 5)
        OR
        (���� = 4 AND ������� = 7)
        OR
        (���� = 5 AND ������� = 9)
        OR
        (���� = 6 AND ������� = 11)
    )
EOT;

$result = $nagruzka->Execute($query);

$flows = array(
    '�' => array(),
    '�' => array(),
    '�' => array()
);
$specs = array();
foreach ($result->GetRows() as $row) {
    $form = $row['������'][0];

    $groupInfo = explode('-', substr($row['������'], 1));
    $flow = $groupInfo[0];

    if (!in_array($row['�������������'], $specs)) {
        $specs[] = $row['�������������'];
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

var_dump(count($flows['�']));
var_dump(count($flows['�']));
var_dump(count($flows['�']));
die();

$subjects = array();
foreach ($result->GetRows() as $row) {
    $subjects[] = $row['����������'];
}

/**
 * ������� ������ ����� �� �������� ���������.
 */
foreach ($subjects as $index => $name) {
    $subjects[$index] = str_replace(
        array(
            ', �/� 1', ', �/� 2', ', �/� 3', ', �/� 4', ', �/� 5', ', �/� 6',
            ', �/� 7', ', �/� 8', ', ���� 1', ', ���� 3', ', ���� 2',
            ', ���� 4', ', ���� 5', ', ���� 6', ', ���� 7', ', ���� 8',
            ', ���� 9', ', ���� 10', ', ����������', ', ��������'
        ),
        '', $name
    );
}
$subjects = array_unique($subjects, SORT_STRING);

/**
 * ��������� � RASSPASS ������ ���������.
 */
$rasspass->Execute('DELETE FROM ����������');
foreach ($subjects as $subject) {
    $rasspass->Execute(
        "INSERT INTO ���������� (����������) VALUES ('{$subject}')"
    );
}

$result = $nagruzka->Execute(
    "SELECT ������, ��������� FROM dbo.�������� WHERE ���������� IN ('���', '���', '��') AND ���������� NOT LIKE '%���������%' AND ��������� > 0"
);
$groups = array();
foreach ($result->GetRows() as $row) {
    if (isset($groups[$row['������']])) {
        if ($groups[$row['������']] != $row['���������']) {
            var_dump($row['������']);
            die();
        }
    } else {
        $groups[$row['������']] = $row['���������'];
    }
}

$specs = array();
$flows = array();

var_dump($groups);
var_dump(count($groups));
