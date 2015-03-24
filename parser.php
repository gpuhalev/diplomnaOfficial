<?php
    require_once 'PHPExcel.php';
    require_once 'PHPExcel/IOFactory.php';

    $myGender = $_REQUEST['myGender'];
    $myEvent = $_REQUEST['myEvent'];
    $myType = $_REQUEST['myType'];
    $myAnnotation = $myType.' for '.$myEvent.' '.$myGender;

    $objReader = PHPExcel_IOFactory::createReader('Excel2007');
    $objPHPExcel = $objReader->load("./uploads/example1.xlsx");
    $worksheet  = $objPHPExcel->setActiveSheetIndexbyName('Sheet1');

    if($myType=="Results"){
        $endingCol = 6;
        $infoCols = $endingCol-3;
    }else{
        $endingCol = 5;
        $infoCols = $endingCol-2;
    }
    $startingRow  = 8;
    
    $highestRow         = $worksheet->getHighestRow();
    $highestColumn      = $worksheet->getHighestColumn();
    $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn);
    $nrColumns          = ord($highestColumn) - 64;
    $worksheetTitle     = $worksheet->getTitle();
?>


<!DOCTYPE html>
<html lang="en">

<head>

    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="description" content="">
    <meta name="author" content="">

    <title><?php echo $myAnnotation?></title>

    <!-- Bootstrap Core CSS -->
    <link href="css/bootstrap.min.css" rel="stylesheet">
    <link href="css/myStyle.css" rel="stylesheet">
    <script type="text/javascript" src="http://code.jquery.com/jquery-1.9.1.min.js"></script>   
</head>

<body>

    <!-- Navigation -->
    <nav class="navbar navbar-inverse navbar-fixed-top" role="navigation">
        <div class="container">
            <div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
                <ul class="nav navbar-nav">
                    <li>
                        <a>
                            <?php
                                echo $myAnnotation;
                            ?>
                        </a>
                    </li>
                </ul>
            </div>
            <!-- /.navbar-collapse -->
        </div>
        <!-- /.container -->
    </nav>

    <!-- Page Content -->
    <div class="container">

        <div class="row">

            <?php

                echo '<table class="table table-striped"><tr>';
                for ($row = $startingRow; $row <= $highestRow; ++ $row) {
                    if($row == $startingRow){
                        echo '<tr style="font-weight:bold;">';
                    }else{
                        echo '<tr>';
                    }
                    for ($col = 0; $col < $endingCol; ++ $col) {
                        $cell = $worksheet->getCellByColumnAndRow($col, $row);
                        $val = $cell->getValue();
                        if($col==0){
                            echo '<td>' . $val .'</td>';
                        }else{
                            echo '<td class="col-md-'.$infoCols.'">' . $val .'</td>';
                        }
                    }
                    echo '</tr>';
                }
                echo '</table>';
                
            ?>

        </div>
        <!-- /.row -->

    </div>
    <!-- /.container -->

    <!-- Bootstrap Core JavaScript -->
    <script src="js/bootstrap.min.js"></script>
    <script src="js/bootstrap-filestyle.min.js"></script>

</body>

</html>
