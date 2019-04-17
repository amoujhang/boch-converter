echo off
for /d %%a in (Input/*) do call :doffmpeg %%a

:doffmpeg
ffmpeg -i "Input/%1/%%05d.png" -loglevel 16 -c:v hap -format hap_alpha -r 30 -vf "scale=trunc(in_w/4)*4:trunc(in_h/4)*4" Output/%1.mov