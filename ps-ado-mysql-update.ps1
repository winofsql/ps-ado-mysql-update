$cn = New-Object -ComObject ADODB.Connection
$rs = New-Object -ComObject ADODB.Recordset

$driver = "{MySQL ODBC 8.0 Unicode Driver}"
$server = "localhost"
$db = "lightbox"
$user = "root"
$pass = ""

$connectionString = "Provider=MSDASQL;Driver={0};Server={1};DATABASE={2};UID={3};PWD={4};"
$connectionString = $connectionString -f $driver,$server,$db,$user,$pass

# 接続文字列表示
$connectionString


try {
    $cn.Open( $connectionString )
}
catch [Exception] {

    $error[0] | Format-List * -force
    exit

}

# 動的カーソル（他ユーザーによる追加・更新・削除を反映）
$rs.CursorType = 2
# レコード単位の共有的ロック（Updateメソッドでのみロックする）
$rs.LockType = 3
$rs.Open( "select * from 社員マスタ where 社員コード <= '0004' ", $cn )

$text = ""
while( !$rs.EOF ) {

    $line = ""

    $line += "{0}{1}" -f $rs.Fields("社員コード").Value, ","
    $line += "{0}{1}" -f $rs.Fields("氏名").Value, ","
    $line += "{0}{1}" -f $rs.Fields("フリガナ").Value, ","
    $line += "{0}{1}" -f $rs.Fields("所属").Value, ","
    $line += "{0}{1}" -f $rs.Fields("性別").Value.ToString(), ","
    $line += "{0}{1}" -f $rs.Fields("給与").Value.ToString(), ","
    $line += "{0}{1}" -f $rs.Fields("手当").Value.ToString(), ","
    $line += "{0}{1}" -f $rs.Fields("管理者").Value, ","
    $line += "{0}{1}" -f $rs.Fields("作成日").Value.ToString("yyyy/MM/dd"), ","
    $line += "{0}{1}" -f $rs.Fields("更新日").Value.ToString("yyyy/MM/dd"), ","
    $line += "{0}{1}" -f $rs.Fields("生年月日").Value.ToString("yyyy/MM/dd"), ","

    $line = $line.Substring(0,$line.Length-1)

    $line += "`n"
    $text += $line

    $rs.Fields("管理者").Value = "0001"
    $rs.Update()

    $rs.MoveNext()
}

# 更新対象のデータ表示
$text

if ( $cn.State -ge 1 ) {
    $cn.Close()
}

