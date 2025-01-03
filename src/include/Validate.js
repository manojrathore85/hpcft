
// cheking email function 

function checkEmail(emailStr)
{
	var pattern = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
	if (emailStr.match(pattern) != null)
			return true;
	else
		return false;
}
//********************** checking the number ************************************

function isNumber(num)
{
	var pattern = /^[\d]+$/;
	if (num.match(pattern) != null)
		return	true;
	else
		return false;
}
	//*********************** checking the name of the person ***************8888888888888888888
	function checkName(name)
	{
			var pattern = /^[A-Za-z][A-Za-z-\s]+[A-Za-z]$/;
			if (name.match(pattern) != null)
				return	true;
			else
				return false;
	}
//******************************************* checking extention ******************************************
function checkExtension(fname)
{
    var extension;

    extension = fname.substring(fname.lastIndexOf('.'),fname.length)

    if ((extension == ".gif") || (extension == ".jpg") || (extension == ".jpeg")  || (extension == ".png") || (extension == ".GIF") || (extension == ".JPG") || (extension == ".JPEG")  || (extension == ".PNG") )
         return true;
    else
           return false;
    
}
//************************************************* checking the maximum and minimum no ****************
function isMaxString(str,min,max)
{
 
    if(str.length < min || str.length > max)
           return true;
    else
    	return false;
 }

//********************************** checking the number with decimal &********************************
function DecimalNumber(num)
 {
 var pattern = /^[\d\.]+$/;
	if (num.match(pattern) != null)
		return	true;
	else
		return false;
 }
 //******************************** chceking the only maximum string
 
		 function MaxString(str,max)
		{
			
			if(str.length > max)
				   return true;
			else
				return false;
		 }
