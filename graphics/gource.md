
# Install the tools

## Gource

Download gource from: https://github.com/acaudwell/Gource/releases

And extract it into a folder like ...
	
    e:\gource-0.42.win32

##  FFMPEG

Get the windows version from here: http://ffmpeg.zeranoe.com/builds/

Store the FFMPEG files in here

    e:\ffmpeg

## Git

If you need to donwload git for Windows get it here: https://git-scm.com/download/win

## Create the video

    set path=%path%;C:\Program Files (x86)\Git\bin
	
    gource.exe E:\code\github\VisioAutomation --seconds-per-day 0.005 --title VisioAutomation --hide filenames,usernames --background 5555dd -viewport 1920x1080 -o d:\visioautomation.ppm

    e:\ffmpeg\ffmpeg.exe -y -r 60 -f image2pipe -vcodec ppm -i d:\visioautomation.ppm -vcodec libx264 -preset ultrafast -crf 1 -threads 0 -bf 0 d:\visioautomation.mp4
