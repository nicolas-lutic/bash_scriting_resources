BEGIN {
last_time=""
}
{
    if ($1 ~ /^[0-9]{4}-[0-9]{2}-[0-9]{2}$/)
    {
        last_time=$1 " " $2
    }
    if (last_time >= start && last_time < end)
    {
        print $0
    }
}
