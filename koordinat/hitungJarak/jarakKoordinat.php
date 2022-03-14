<?php
/**
 * Menghitung jarak dari 2 koordinat
 */

function jarak($position1, $position2 , $unit = null)
{
    $loc1 = explode(",",$position1);
    $loc2 = explode(",",$position2);
    $lat1 = $loc1[0]; $lon1 = $loc1[1];
    $lat2 = $loc2[0]; $lon2 = $loc2[1];
    $theta = $lon1 - $lon2;
    $dist = sin(deg2rad($lat1)) * sin(deg2rad($lat2)) +  cos(deg2rad($lat1)) * cos(deg2rad($lat2)) * cos(deg2rad($theta));
    $dist = acos($dist);
    $dist = rad2deg($dist);
    $miles = $dist * 60 * 1.1515;
    $unit = strtoupper($unit);

    if ($unit == "M") {
        $jarak = $miles;
    } else if ($unit == "N") {
        $jarak = ($miles * 0.8684);
    } else {
        //jarak dalam meter
        $jarak = ($miles * 1609.33999997549);
        
        //jika jarak lebih dari 1000 meter akan di ubah menjadi km
        if($jarak>1000){
            $jarak = $jarak/1000;
        }
    }
    
    return $jarak;
}
