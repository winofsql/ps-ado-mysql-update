$cn = New-Object -ComObject ADODB.Connection
$rs = New-Object -ComObject ADODB.Recordset

$driver = "{MySQL ODBC 8.0 Unicode Driver}"
$server = "localhost"
$db = "lightbox"
$user = "root"
$pass = ""

$connectionString = "Provider=MSDASQL;Driver={0};Server={1};DATABASE={2};UID={3};PWD={4};"
$connectionString = $connectionString -f $driver,$server,$db,$user,$pass

# �ڑ�������\��
$connectionString


try {
    $cn.Open( $connectionString )
}
catch [Exception] {

    $error[0] | Format-List * -force
    exit

}

# ���I�J�[�\���i�����[�U�[�ɂ��ǉ��E�X�V�E�폜�𔽉f�j
$rs.CursorType = 2
# ���R�[�h�P�ʂ̋��L�I���b�N�iUpdate���\�b�h�ł̂݃��b�N����j
$rs.LockType = 3
$rs.Open( "select * from �Ј��}�X�^ where �Ј��R�[�h <= '0004' ", $cn )

$text = ""
while( !$rs.EOF ) {

    $line = ""

    $line += "{0}{1}" -f $rs.Fields("�Ј��R�[�h").Value, ","
    $line += "{0}{1}" -f $rs.Fields("����").Value, ","
    $line += "{0}{1}" -f $rs.Fields("�t���K�i").Value, ","
    $line += "{0}{1}" -f $rs.Fields("����").Value, ","
    $line += "{0}{1}" -f $rs.Fields("����").Value.ToString(), ","
    $line += "{0}{1}" -f $rs.Fields("���^").Value.ToString(), ","
    $line += "{0}{1}" -f $rs.Fields("�蓖").Value.ToString(), ","
    $line += "{0}{1}" -f $rs.Fields("�Ǘ���").Value, ","
    $line += "{0}{1}" -f $rs.Fields("�쐬��").Value.ToString("yyyy/MM/dd"), ","
    $line += "{0}{1}" -f $rs.Fields("�X�V��").Value.ToString("yyyy/MM/dd"), ","
    $line += "{0}{1}" -f $rs.Fields("���N����").Value.ToString("yyyy/MM/dd"), ","

    $line = $line.Substring(0,$line.Length-1)

    $line += "`n"
    $text += $line

    $rs.Fields("�Ǘ���").Value = "0001"
    $rs.Update()

    $rs.MoveNext()
}

# �X�V�Ώۂ̃f�[�^�\��
$text

if ( $cn.State -ge 1 ) {
    $cn.Close()
}

