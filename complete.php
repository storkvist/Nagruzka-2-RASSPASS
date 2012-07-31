<?php

date_default_timezone_set('Europe/Moscow');

include 'tcpdf/tcpdf.php';
include 'adodb517/adodb.inc.php';

/**
 * Загрузка данных из CSV файла с табуляцией в качестве разделителя.
 *
 * @param $filename
 * @return array
 */
function loadCsvData($filename) {
    $result = array();

    $handle = fopen($filename, 'r');
    if ($handle) {
        while ($str = fgets($handle)) {
            $result[] = explode("\t", $str);
        }
    }

    return $result;
}

/**
 * Аналог функции explode() для UTF-8 текста.
 *
 * @param $delimiter
 * @param $string
 * @param $limit
 * @param string $encoding
 * @return array
 */
function mb_explode($delimiter, $string, $limit = -1, $encoding = 'auto') {
    if(!is_array($delimiter)) {
        $delimiter = array($delimiter);
    }
    if(strtolower($encoding) === 'auto') {
        $encoding = mb_internal_encoding();
    }
    if(is_array($string) || $string instanceof Traversable) {
        $result = array();
        foreach($string as $key => $val) {
            $result[$key] = mb_explode($delimiter, $val, $limit, $encoding);
        }
        return $result;
    }

    $result = array();
    $currentpos = 0;
    $string_length = mb_strlen($string, $encoding);
    while($limit < 0 || count($result) < $limit) {
        $minpos = $string_length;
        $delim_index = null;
        foreach($delimiter as $index => $delim) {
            if(($findpos = mb_strpos($string, $delim, $currentpos, $encoding))
                !== false
            ) {
                if($findpos < $minpos) {
                    $minpos = $findpos;
                    $delim_index = $index;
                }
            }
        }
        $result[] = mb_substr(
            $string, $currentpos, $minpos - $currentpos, $encoding
        );
        if($delim_index === null) {
            break;
        }
        $currentpos = $minpos + mb_strlen($delimiter[$delim_index], $encoding);
    }
    return $result;
}

function u2w($text) {
    return mb_convert_encoding($text, 'windows-1251', 'UTF-8');
}

function w2u($text) {
    return mb_convert_encoding($text, 'UTF-8', 'windows-1251');
}

/**
 * Получаем данные из вычищенных Excel файлов.
 */
if (mb_strstr($_SERVER['PATH'], 'C:\WINDOWS')) {
    $envPath = 'Z:\Sites\Apps\NAGRUZKA-2-RASSPASS\\';
    header('Content-Type: text/html; charset=utf-8');
    $data = array_merge(
        loadCsvData('Z:\Sites\Apps\NAGRUZKA-2-RASSPASS\nagruzkaUTF8.txt'),
        loadCsvData('Z:\Sites\Apps\NAGRUZKA-2-RASSPASS\nagruzkaVUTF8.txt')
    );
} else {
    $envPath = '';
    $data = array_merge(
        loadCsvData('nagruzkaUTF8.txt'),
        loadCsvData('nagruzkaVUTF8.txt')
    );
}

$wrongEncodingData = $data;
$data = array();
foreach ($wrongEncodingData as $index => $row) {
    foreach ($row as $index2 => $field) {
        $data[$index][$index2] = trim(mb_convert_encoding(
            mb_convert_encoding($field, 'windows-1251', 'UTF-8'),
            'UTF-8',
            'windows-1251'
        ));
    }
}

/**
 * Определяем уникальные предметы.
 */
$uniqueSubjects = array();
foreach ($data as $row) {
    $subject = str_replace(array(
        ', п/г 1', ', п/г 2', ', п/г 3', ', п/г 4', ', п/г 5', ', п/г 6',
        ', преп 1', ', преп 2', ', часть 1', ', часть 2', ', часть_к 1',
        ', часть_к 2'
    ), '', $row[0]);
    $uniqueSubjects[] = $subject;
}
$uniqueSubjects = array_unique($uniqueSubjects);

$uniqueSpecs = array();
$sport = array();
$SPECS = array(
    'Д' => array(),
    'В' => array()
);
$FLOWS = array(
    'Д' => array(),
    'В' => array()
);
foreach ($data as $row) {
    /**
     * Выносим практику по физкультуре в отдельные данные — с ней лучше
     * разбираться индивидуально.
     */
    if (mb_strstr($row[0], 'Физическая культура')) {
        $sport[] = $row;
        continue;
    }

    /**
     * Убираем нахер второе высшее.
     */
    if (mb_strstr($row[2], '2во')) {
        continue;
    }

    /**
     * Раскурочиваем название группы на форму обучения, специальность, курс
     * и группу.
     */

    /**
     * @var string $form Форма обучения (Д/В).
     */
    $form = mb_substr($row[2], 0, 1, 'UTF-8');

    /**
     * @var string $group Название группы (Цим-2-1).
     */
    $group = mb_substr($row[2], 2);

    $groupInfo = mb_explode('-', $group, -1, 'UTF-8');

    /**
     * @var string $spec Специальность группы (Цим).
     */
    $spec = $groupInfo[0];

    /**
     * @var int $course Курс группы (2).
     */
    $course = intval($groupInfo[1]);

    /**
     * @var int $subgroup Номер подгруппы.
     */
    $subgroup = intval($groupInfo[2]);

    if (isset($SPECS[$form][$spec])) {
        if (isset($SPECS[$form][$spec][$course])) {
            if (isset($SPECS[$form][$spec][$course][$subgroup])) {
                $SPECS[$form][$spec][$course][$subgroup][] = $row;
                $FLOWS[$form][$spec][$course][$subgroup] = max(
                    $FLOWS[$form][$spec][$course][$subgroup],
                    intval($row[3])
                );
            } else {
                $SPECS[$form][$spec][$course][$subgroup] = array($row);
                $FLOWS[$form][$spec][$course][$subgroup] = intval($row[3]);
            }
        } else {
            $SPECS[$form][$spec][$course] = array(
                $subgroup => array($row)
            );
            $FLOWS[$form][$spec][$course] = array(
                $subgroup => intval($row[3])
            );
        }
    } else {
        $SPECS[$form][$spec] = array(
            $course => array(
                $subgroup => array($row)
            )
        );
        $FLOWS[$form][$spec] = array(
            $course => array(
                $subgroup => intval($row[3])
            )
        );
    }

    /**
     * Определяем уникальные специальности.
     */
    $uniqueSpecs[] = $spec;
}
$uniqueSpecs = array_unique($uniqueSpecs);

/**
 * Сортируем всё по ключам.
 */
foreach ($SPECS as $form => $spec) {
    ksort($spec);
    foreach ($spec as $specName => $course) {
        ksort($course);
        foreach ($course as $courseNumber => $subgroup) {
            ksort($subgroup);
        }
    }
}

/**
 * Получившаяся структура массива $SPECS:
 * -> Форма обучения
 *     -> Специальность
 *         -> Курс
 *             -> Группа
 *                 -> Дисциплины
 */

/**
 * Схлопываем предметы в рамках отдельных групп.
 */
foreach ($SPECS as $form => $spec) {
    foreach ($spec as $specName => $course) {
        foreach ($course as $courseNumber => $group) {
            foreach ($group as $groupNumber => $rows) {
                /**
                 * Занятия сгруппированные по кафедрам и предметам.
                 */
                $lecSubjects = array();

                /**
                 * Выносим в отдельную группу лекции.
                 */
                foreach ($rows as $index => $row) {
                    if ('Лек' === $row[5]) {
                        $subjectName = str_replace(array(
                            ', п/г 1', ', п/г 2', ', п/г 3', ', п/г 4',
                            ', п/г 5', ', п/г 6', ', преп 1', ', преп 2',
                            ', часть 1', ', часть 2', ', часть_к 1',
                            ', часть_к 2',
                        ), '', $row[0]);
                        $department = intval($row[11]);
                        if (isset($lecSubjects[$department])) {
                            $lecSubjects[$department][$subjectName][] = $row;
                        } else {
                            $lecSubjects[$department] = array(
                                $subjectName => array($row)
                            );
                        }
                    }
                }

                /**
                 * Проходимся по практике и лабораторным работам.
                 */
                foreach ($rows as $index => $row) {
                    if ('Лаб' === $row[5] || 'Пр' === $row[5]) {
                        $subjectName = str_replace(array(
                            ', п/г 1', ', п/г 2', ', п/г 3', ', п/г 4',
                            ', п/г 5', ', п/г 6', ', преп 1', ', преп 2',
                            ', часть 1', ', часть 2', ', часть_к 1',
                            ', часть_к 2',
                        ), '', $row[0]);
                        $department = intval($row[11]);
                        if (isset($lecSubjects[$department][$subjectName])) {
                            $lecSubjects[$department][$subjectName][] = $row;
                        } else {
                            if (isset($lecSubjects[$department])) {
                                if (isset($lecSubjects[$department][$subjectName])) {
                                    $lecSubjects[$department][$subjectName][] = $row;
                                } else {
                                    $lecSubjects[$department][$subjectName] = array($row);
                                }
                            } else {
                                $lecSubjects[$department] = array(
                                    $subjectName => array($row)
                                );
                            }
                        }
                    }
                }

                $SPECS[$form][$specName][$courseNumber][$groupNumber] = $lecSubjects;
            }
        }
    }
}

/**
 * Тепеь схлопываем предметы в рамках потоков
 * (Форма обучения + Специальность + Курс).
 *
 * Находим несовпадения в часах между группами.
 */
$fuckedPlan = array();
$RESPECS = array();
foreach ($SPECS as $form => $spec) {
    foreach ($spec as $specName => $course) {
        foreach ($course as $courseNumber => $group) {
            $flow = array();
            foreach ($group as $groupNumber => $groups) {
                if (!isset($flow[$groupNumber])) {
                    $flow[$groupNumber] = array();
                }

                foreach ($groups as $departmentNumber => $subjects) {
                    if (!isset($flow[$groupNumber][$departmentNumber])) {
                        $flow[$groupNumber][$departmentNumber] = array();
                    }

                    foreach ($subjects as $subjectName => $rows) {
                        if (!isset($flow[$groupNumber][$departmentNumber][$subjectName])) {
                            $flow[$groupNumber][$departmentNumber][$subjectName] = array();
                        }

                        /**
                         * Часы по дисциплине
                         */
                        $flow[$groupNumber][$departmentNumber][$subjectName] = array(
                            'baseHours' => array(
                                'lecHour'   => 0,
                                'praHour'   => 0,
                                'labHour'   => 0
                            )
                        );
                        $lecHour = 0;
                        $praHour = 0;
                        $labHour = 0;

                        foreach ($rows as $row) {
                            if ('Лек' === $row[5]) {
                                $flow[$groupNumber][$departmentNumber][$subjectName]
                                    ['baseHours']['lecHour'] = intval($row[6]);
                            } elseif ('Пр' === $row[5]) {
                                $flow[$groupNumber][$departmentNumber][$subjectName]
                                    ['baseHours']['praHour']
                                        = intval($row[6]);
                            } elseif ('Лаб' === $row[5]) {
                                $flow[$groupNumber][$departmentNumber][$subjectName]
                                    ['baseHours']['labHour']
                                        = intval($row[6]);
                            }

                            $flow[$groupNumber][$departmentNumber][$subjectName][] = $row;
                        }
                    }
                }
            }

            $reFlow = array();
            foreach ($flow as $groupNumber => $groups) {
                foreach ($groups as $departmentNumber => $subjects) {
                    if (!isset($reFlow[$departmentNumber])) {
                        $reFlow[$departmentNumber] = array();
                    }

                    foreach ($subjects as $subjectName => $rows) {
                        if (!isset($reFlow[$departmentNumber][$subjectName])) {
                            $reFlow[$departmentNumber][$subjectName] = array();
                            $reFlow[$departmentNumber][$subjectName]['baseHours'] = $rows['baseHours'];
                        } else {
                            /**
                             * Проверяем, что в рамках дисциплины совпадают часы
                             * по группам.
                             */
                            $flag = true;
                            $flag = $flag && ($reFlow[$departmentNumber][$subjectName]['baseHours']['lecHour'] === $rows['baseHours']['lecHour']);
                            $flag = $flag && ($reFlow[$departmentNumber][$subjectName]['baseHours']['praHour'] === $rows['baseHours']['praHour']);
                            $flag = $flag && ($reFlow[$departmentNumber][$subjectName]['baseHours']['labHour'] === $rows['baseHours']['labHour']);
                            if (!$flag) {
                                $fuckedPlan[] = array(
                                    'form'          => $form,
                                    'spec'          => $specName,
                                    'course'        => $courseNumber,
                                    'department'    => $departmentNumber,
                                    'subject'       => $subjectName
                                );
                            }
                        }

                        foreach ($rows as $index => $row) {
                            if ('baseHours' !== $index) {
                                if (isset($reFlow[$departmentNumber][$subjectName][$groupNumber])) {
                                    $reFlow[$departmentNumber][$subjectName][$groupNumber][] = $row;
                                } else {
                                    $reFlow[$departmentNumber][$subjectName][$groupNumber] = array($row);
                                }
                            }
                        }
                    }
                }
            }

            $SPECS[$form][$specName][$courseNumber] = $reFlow;
        }
    }
}

/**
 * На этом этапе структура $SPECS такова:
 * -> Форма обучения
 *     -> Специальность
 *         -> Курс
 *             -> Кафедра
 *                 -> Дисциплина
 *                     -> baseHours
 *                         -> lecHour
 *                         -> praHour
 *                         -> labHour
 *                     -> Циклы
 */

/**
 * Убираем из наших данных странные планы — с ними надо разбираться вручную.
 */
foreach ($fuckedPlan as $row) {
    unset($SPECS[$row['form']][$row['spec']][$row['course']][$row['department']][$row['subject']]);
}

/**
 * Путь к базе данных РАСПАСС.
 */
$RASSPASS = 'C:\\2012_2013 1\\Raspis1.mdb';

/**
 * @var ADOConnection $rasspass
 */
$rasspass = ADONewConnection('access');
//$rasspass->debug = true;
$rasspass->SetFetchMode(ADODB_FETCH_ASSOC);
$dsn = "Driver={Microsoft Access Driver (*.mdb)};Dbq={$RASSPASS};Uid=Admin;";
$rasspass->Connect($dsn);
//$rasspass->debug = true;

$rasspass->Execute(u2w('DELETE * FROM Планы'));
$rasspass->Execute(u2w('DELETE * FROM Дисциплины'));
$rasspass->Execute(u2w('DELETE * FROM Потоки'));
$rasspass->Execute(u2w('DELETE * FROM Специальности'));
$rasspass->Execute(u2w('DELETE * FROM Кафедры'));


/**
 * Коды кафедр.
 */
$_NAGRUZKA_DEPARTMENTS = array(
    1   => 'Физ',
    2   => 'Хим',
    3   => 'Мат',
    4   => 'ТДП',
    5   => 'ТПиПП',
    6   => 'ТиТЦП',
    7   => 'ИТиУ',
    8   => 'ПиПО',
    9   => 'АПП',
    10   => 'МСиТ',
    11   => 'ИС',
    12   => 'ИиВТ',
    13   => 'НГ',
    14   => 'ВМ',
    15   => 'Англ',
    16   => 'ФУиА',
    17   => 'УБП',
    18   => 'Мндж',
    19   => 'ПМиМС',
    20   => 'ФВС',
    21   => 'ЭиИ',
    22   => 'ИДиР',
    23   => 'КБ',
    24   => 'КиПК',
    25   => 'ИЛ',
    26   => 'РЯиС',
    27   => 'ПП',
    28   => 'МКиП',
    29   => 'РМ',
    30   => 'УРБ',
    31   => 'СсО',
    33   => 'Ист',
    35   => 'ИллЭст',
    36   => 'РиЖ',
    37   => 'ХТОПП',
    38   => 'Фил',
    39   => 'ПиС',
    40   => 'ИКАБТ',
);
foreach ($_NAGRUZKA_DEPARTMENTS as $departmentName) {
    $rasspass->Execute(u2w("INSERT INTO Кафедры (Кафедра) VALUES ('{$departmentName}');"));
}

$_DEPARTMENTS = array();
$rows = $rasspass->Execute(u2w('SELECT КодКафедры, Кафедра FROM Кафедры'));
foreach ($rows->GetRows() as $row) {
    $_DEPARTMENTS[w2u($row[u2w('Кафедра')])] = intval($row[u2w('КодКафедры')]);
}

/**
 * Заносим дисицплины в РАСПАСС.
 */
foreach ($uniqueSubjects as $subject) {
    $rasspass->Execute(
        u2w("INSERT INTO Дисциплины (Дисциплина) VALUES ('{$subject}')")
    );
}

/**
 * Коды дисциплин.
 */
$_SUBJECTS = array();
$rows = $rasspass->Execute(u2w('SELECT Код, Дисциплина FROM Дисциплины'));
foreach ($rows->GetRows() as $row) {
    $_SUBJECTS[w2u($row[u2w('Дисциплина')])] = $row[u2w('Код')];
}

/**
 * Заполняем таблицы Специальности и Потоки.
 */
foreach ($FLOWS as $form => $spec) {
    foreach ($spec as $specName => $groups) {
        $rasspass->Execute(
            u2w("INSERT INTO Специальности (Спец) VALUES ('{$specName}')")
        );
        foreach ($groups as $course => $theGroups) {
            $groupCount = 0;
            $studentCount = 0;
            foreach ($theGroups as $stCount) {
                $groupCount++;
                $studentCount += $stCount;

            }
            $query = <<<EOT
INSERT INTO Потоки
    (ФормаОбучения, Спец, Курс, Групп, Студентов, Начало, Конец, КолРабочихДней)
VALUES (
    '{$form}', '{$specName}', '{$course}', {$groupCount}, {$studentCount},
    #09/01/2012#, #12/29/2012#, 5
)
EOT;
            $rasspass->Execute(u2w($query));
        }
    }
}

/**
 * Получаем коды вставленных потоков.
 */
$_FLOWS = array(
    'В' => array(),
    'Д' => array()
);
$rows = $rasspass->Execute(u2w('SELECT КодПотока, Спец, Курс, ФормаОбучения FROM Потоки'));
foreach ($rows->GetRows() as $row) {
    if (isset($_FLOWS[w2u($row[u2w('ФормаОбучения')])][w2u($row[u2w('Спец')])])) {
        $_FLOWS[w2u($row[u2w('ФормаОбучения')])][w2u($row[u2w('Спец')])][intval(w2u($row[u2w('Курс')]))] = intval(w2u($row[u2w('КодПотока')]));
    } else {
        $_FLOWS[w2u($row[u2w('ФормаОбучения')])][w2u($row[u2w('Спец')])] = array(
            intval(w2u($row[u2w('Курс')]))
                => intval(w2u($row[u2w('КодПотока')]))
        );
    }
}

$rasspass->Disconnect();
$rasspass->Connect($dsn);

/**
 * Заполняем таблицу с планами.
 */
//$rasspass->debug = true;
foreach ($SPECS as $form => $flow) {
    foreach ($flow as $specName => $courses) {
        foreach ($courses as $courseNumber => $departments) {
            foreach ($departments as $departmentNumber => $subjects) {
                foreach ($subjects as $subjectName => $rows) {
                    if ('' == $_DEPARTMENTS[$_NAGRUZKA_DEPARTMENTS[intval($departmentNumber)]]) {
                        var_dump($rows);
                        echo '<hr>';
                    }

                    if ('' == $_SUBJECTS[$subjectName]) {
                        var_dump($rows);
                        echo '<hr>';
                    }


                    $query = <<<EOT
INSERT INTO Планы (КодПотока, КодКафедры, Дисциплины_Код,
                   Лекций, Семинаров, Лабораторных)
VALUES ({$_FLOWS[$form][$specName][$courseNumber]},
        {$_DEPARTMENTS[$_NAGRUZKA_DEPARTMENTS[intval($departmentNumber)]]},
        {$_SUBJECTS[$subjectName]}, {$rows['baseHours']['lecHour']},
        {$rows['baseHours']['praHour']}, {$rows['baseHours']['labHour']});
EOT;
                    $rasspass->Execute(u2w($query));
                }
            }
        }
    }
}

/**
 * Печатаем учебные планы с физкультурой.
 */
$sportPdf = new TCPDF();
$sportPdf->addTTFfont($envPath . 'PTSans/PTS55F.ttf');
$sportPdf->addTTFfont($envPath . 'PTSans/PTS75F.ttf');
$sportPdf->addTTFfont($envPath . 'PTSans/PTC55F.ttf');
$sportPdf->SetFont('pts55f', '', 10, true, false);
$sportPdf->setPageOrientation('Landscape');
$sportPdf->setPrintHeader(false);
$sportPdf->setPrintFooter(false);

$xhtml = <<<EOT
<h1>Учебные планы по физической культуре</h1>
<table border="1" cellpadding="5">
    <tr>
        <td>Группа</td>
        <td>Студентов</td>
        <td>Тип</td>
        <td>Часов</td>
        <td>Преподаватель</td>
        <td>Поток</td>
        <td>И</td>
    </tr>
EOT;

$zebra = false;
foreach ($sport as $row) {
    if ($zebra) {
        $style="background-color: #BBBBBB;";
    } else {
        $style = '';
    }
    $zebra = !$zebra;

    $pg = '';
    if (mb_strstr($row[0], 'п/г 1')) {
        $pg = ', п/г 1';
    }
    if (mb_strstr($row[0], 'п/г 2')) {
        $pg = ', п/г 2';
    }

    $xhtml .= <<<EOT
    <tr style="{$style}">
        <td>{$row[2]}{$pg}</td>
        <td>{$row[3]}</td>
        <td>{$row[5]}</td>
        <td>{$row[6]}</td>
        <td>{$row[7]}</td>
        <td>{$row[8]}</td>
        <td>{$row[9]}</td>
    </tr>
EOT;
}

$xhtml .= '</table>';
$sportPdf->AddPage();
$sportPdf->writeHTML($xhtml);
$sportPdf->Output($envPath . 'sportPdf.pdf', 'F');


/**
 * Печатаем учебные планы с дурацкими предметами.
 */
$fuckedPdf = new TCPDF();
$fuckedPdf->addTTFfont($envPath . 'PTSans/PTS55F.ttf');
$fuckedPdf->addTTFfont($envPath . 'PTSans/PTS75F.ttf');
$fuckedPdf->addTTFfont($envPath . 'PTSans/PTC55F.ttf');
$fuckedPdf->SetFont('pts55f', '', 10, true, false);
$fuckedPdf->setPageOrientation('Landscape');
$fuckedPdf->setPrintHeader(false);
$fuckedPdf->setPrintFooter(false);

$xhtml = <<<EOT
<h1>Учебные планы с очень странной разбивкой по часам</h1>
<table border="1" cellpadding="5">
    <tr>
        <td>Дициплина</td>
        <td>Группа</td>
        <td>Студентов</td>
        <td>Тип</td>
        <td>Часов</td>
        <td>Преподаватель</td>
        <td>Поток</td>
        <td>И</td>
    </tr>
EOT;

$zebra = false;
foreach ($fuckedPlan as $plan) {
    foreach ($data as $row) {
        if (
            mb_strstr($row[2], $plan['form'] . $plan['spec'] . '-' . $plan['course'])
            &&
            mb_strstr($row[0], $plan['subject'])
        ) {
            if ($zebra) {
                $style="background-color: #BBBBBB;";
            } else {
                $style = '';
            }
            $zebra = !$zebra;

            $xhtml .= <<<EOT
    <tr style="{$style}">
        <td>{$row[0]}</td>
        <td>{$row[2]}</td>
        <td>{$row[3]}</td>
        <td>{$row[5]}</td>
        <td>{$row[6]}</td>
        <td>{$row[7]}</td>
        <td>{$row[8]}</td>
        <td>{$row[9]}</td>
    </tr>
EOT;
        }
    }
}

$xhtml .= '</table>';
$fuckedPdf->AddPage();
$fuckedPdf->writeHTML($xhtml);
$fuckedPdf->Output($envPath . 'fuckedPdf.pdf', 'F');
