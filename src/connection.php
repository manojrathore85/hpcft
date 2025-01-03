<?php session_start();
//error_reporting(0);
  error_reporting( E_ALL );
//date_default_timezone_set("Asia/Kolkata");
//date_default_timezone_set("Australia/NSW"); //CR20072020
$date = date('Y-m-d');
$datetime = date('Y-m-d H:i:s');
$ip = $_SERVER['REMOTE_ADDR'];

require_once ("mode.php");

class DB extends MODE{	

	protected	$hostname_server="localhost";
	protected	$username_server="root";
	protected	$password_server="";
	protected	$database_server="cims_ims";

	// protected	$hostname_server="localhost";
	// protected	$username_server="imstestclientusr";
	// protected	$password_server="?u410qIc3";
	// protected	$database_server="ims_testclient";
/*
	protected	$hostname_server="localhost";
	protected	$username_server="essennyl_tstusr";
	protected	$password_server="_}^xZp-5GFhB";
	protected	$database_server="essennyl_test";
		*/
	public function connect() {
  
 	  if($this->getmode()=='local')
	  {
		 $db = mysqli_connect($this->hostname_local, $this->username_local, $this->password_local,$this->database_local); 		
	  }else
	  {	 	 
		 $db = mysqli_connect($this->hostname_server, $this->username_server, $this->password_server,$this->database_server);		
	  }	 
		return $db;
	}
	 
	public function dateFormat($date)
	{
		 if($date !="" && $date != "0000-00-00 00:00:00")
		 {
		 $dateStr = strtotime($date);
		 $newDate = date('d-m-Y H:i',$dateStr);
		 }
		 else
		 {
		 $newDate = "";	 
		 }
	 return  $newDate;	
	}
	public function bookingStatus($status)
	{
		 $tagline='';
	//	if($status=='Approved'){$tagline='<span class="label label-primary">Approved</span>';}
	//	elseif($status=='Rejected'){$tagline= '<span class="label label-danger">Rejected</span>';}
		if($status=='Cancelled'){$tagline= '<span class="label label-danger">Cancelled</span>';}
		elseif($status=='Assigned'){$tagline= '<span class="label label-primary">Assigned</span>';}
		elseif($status=='Completed'){$tagline= '<span class="label label-success">Completed</span>';}
		elseif($status=='In Progress'){$tagline= '<span class="label bg-yellow">In Progress</span>';}
		else{$tagline= '<span class="label label-warning">Pending</span>';}
		
		return $tagline;
	}
    public function bookingColor($status)
	{
		$tagline='#ECF0F5'; //Red by default
		
	//	if($status=='Approved'){$tagline='blue';}
	//	elseif($status=='Rejected'){$tagline= '#cc3a3a';}
		if($status=='Cancelled'){$tagline= 'red';}
		elseif($status=='Assigned'){$tagline= '#367FA9';}
		elseif($status=='Completed'){$tagline= 'green';}
		elseif($status=='In Progress'){$tagline= '#e2e230';}
		else{$tagline= 'purple';}
		
		return $tagline;
		
	}
	public function Redirect($url)
	{
		echo "<span style='padding-left:600px;font-size:20px'><strong>Please wait..</strong></span>"; 
		echo '<script>document.location.href="'.$url.'";</script>';
		exit();
	}
	public function sendMail($email,$subject,$mailbody,$attachment='')
	{
	//return true;
	global  $datetime,$ip;
		$db=$this->connect();
		$subject1=mysqli_real_escape_string($db,$subject);
		$mailbody1=mysqli_real_escape_string($db,$mailbody);
	 $query_insert="insert into `email_log`(`email`, `subject`, `mailbody`, `senddate`, `attachment`) values('$email','$subject1','$mailbody1','$datetime','$attachment') ";	
		   mysqli_query($db,$query_insert);
		   $id = mysqli_insert_id($db);	
	
	include_once("class.phpmailer.php");			
			
			$mail = new PHPMailer();
			$mail->SMTPSecure    = 'ssl';
			$mail->SMTPDebug=false;
			$mail->Host = 'mail.essential4transport.com.au';
			$mail->Port = 465;
			$mail->IsSMTP();
			$mail->SMTPAuth = true;
			$mail->From = 'admin@essential4transport.com.au';
			$mail->FromName = "Essential4Transport";//$sch_nm;
			$mail->Username = "system@essential4transport.com.au";
			$mail->Password = "qrRmXO^wve7G";
			
			$mail->AddAddress($email,'User');
		//	$mail->AddAddress('prajapati.sanju5@gmail.com','Sanjay');
			
			$mail->AddBCC('agrawalatul@gmail.com','Atul');
			
			if(trim($attachment)!='')
			{
				$filepath = $_SERVER["DOCUMENT_ROOT"].'/receipts/'.$attachment;					
				if(file_exists($filepath))
				{
				$mail->AddAttachment($filepath,'Invoice.pdf');
				}
			}
			
			$mail->IsHTML(true);
			$mail->Subject = $subject;
			$mail->Body = $mailbody;
			
			if ($mail->Send())
				{
					$update="update `email_log` set `status`=1 where `id`=$id";
					mysqli_query($db,$update);
					return true;
				}
			else
				{
					return false;
				}
	
	}	 
}
?>