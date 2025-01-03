<?php
require_once("connection.php");
$object = new DB();
$con =$object->connect();
$prjid = $_REQUEST['projectid'];
$issueid = $_REQUEST['issueid'];
$comments = mysqli_real_escape_string($con,$_REQUEST['txtcomments']);
$updatedate = date("Y-m-d H:i:s"); 
$user = $_REQUEST['userid'];

$query = "insert into issuechangescomments (`projectid`,`issueid`,`comments`,`updateDate`,`changesby`) 
values('$prjid','$issueid','$comments','$updatedate','$user')";
if (mysqli_set_charset($con,"utf8")) {
	mysqli_query($con,$query);
}

echo "Done";

?>
