/*! Copyright (c) 2020 Zick Keen (https://github.com/zickkeen)
 * Licensed Under MIT (http://opensource.org/licenses/MIT)
 *
 * Version 0.0.1
 *
 * 
 */



/***
 * Convert Coordinate to Degrees Minutes Seconds (DMS)
 */
function toDegrees(coordinate) {
    var absolute = Math.abs(coordinate);
    var degrees = Math.floor(absolute);
    var minutesNotTruncated = (absolute - degrees) * 60;
    var minutes = Math.floor(minutesNotTruncated);
    var seconds = Math.floor((minutesNotTruncated - minutes) * 60);

    return degrees + "°" + minutes + "'" + seconds + "\"";
}

function toDms(coordinate) {
    var kor = coordinate.split(",");
    var latitude = toDegrees(kor[0]);
    var latitudeCardinal = kor[0] >= 0 ? "N" : "S";

    var longitude = toDegrees(kor[1]);
    var longitudeCardinal = kor[1] >= 0 ? "E" : "W";

    return latitude + latitudeCardinal + "+" + longitude + longitudeCardinal;
}

function urlLocation(dms, coordinate) {
    var koor = coordinate.replace(/\s/g, '');
    var url = "https://www.google.com/maps/place/" + dms + "/" + koor + "/";
    return url;
}