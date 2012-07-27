<?php

$RASSPASS = 'Z:\\Sites\\Apps\\Nagruzka-2-RASSPASS\\test.mdb';

include 'adodb517/adodb.inc.php';

/**
 * @var ADOConnection $nagruzka
 */
$nagruzka = ADONewConnection('odbc_mssql');
$nagruzka->debug = true;
$nagruzka->SetFetchMode(ADODB_FETCH_ASSOC);

$dsn = 'DRIVER={SQL Server};SERVER=VITALY-XP\SQLEXPRESS;DATABASE=Деканат;';
$nagruzka->Connect($dsn);

$result = $nagruzka->Execute('SELECT * FROM dbo.dmDisc');
var_dump($result->GetRows());

/**
 * @var ADOConnection $rasspass
 */
$rasspass = ADONewConnection('access');
$rasspass->debug = true;
$rasspass->SetFetchMode(ADODB_FETCH_ASSOC);


$dsn = "Driver={Microsoft Access Driver (*.mdb)};Dbq={$RASSPASS};Uid=Admin;";
$rasspass->Connect($dsn);

$rs = $rasspass->Execute('SELECT * FROM Доставка');
var_dump($rs->GetRows());
