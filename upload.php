<?php
$response = array(
	'status'=>0,
	'message'=>''
);
if(isset($_FILES['testfile']) && @$_FILES["testfile"]["error"] == 0)
{
	$response['status'] = 1;
	$response['message'] = 'Success!';
}
elseif(@$_FILES["testfile"]["error"])
{
    $response['message'] = 'Error code: '.$_FILES["testfile"]["error"].'.';
}
elseif ($_SERVER['REQUEST_METHOD'] == 'POST' && empty($_POST) && $_SERVER['CONTENT_LENGTH'] > 0)
{
    $response['message'] = sprintf('The server was unable to handle that much POST data (%s bytes) due to its current configuration', $_SERVER['CONTENT_LENGTH']).'.';
}
else
{
    $response['message'] = 'Unknown error.';
}
echo json_encode($response);
?>