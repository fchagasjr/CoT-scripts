$path = $Args[0]

If (!(test-path $path))
{
    mkdir $path
}

cd $path