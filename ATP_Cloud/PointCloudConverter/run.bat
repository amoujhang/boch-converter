SET INPUT=%~1
SET OUTPUT=%~2
call LAStools\bin\e572las.exe -v -i "%INPUT%" -o tmp\out.laz &
call PotreeConverter.exe tmp\out.laz --overwrite -o "%OUTPUT%" &