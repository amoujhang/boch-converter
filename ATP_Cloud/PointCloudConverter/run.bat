SET INPUT=%~1
SET OUTPUT=%~nx1
mkdir "output\%OUTPUT%" &
call LAStools\bin\e572las.exe -v -i "%INPUT%" -o tmp\out.laz &
call PotreeConverter.exe tmp\out.laz --overwrite -o "output\%OUTPUT%" &
call "C:\Program Files\7-Zip\7z.exe" a -t7z "output\%OUTPUT%.zip" "output\%OUTPUT%"