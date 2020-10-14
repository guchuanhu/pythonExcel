from datetime import datetime, timedelta, timezone
def getTime():
    SHA_TZ = timezone(
        timedelta(hours=8),
        name='Asia/Shanghai',
    )
    utc_now = datetime.utcnow().replace(tzinfo=timezone.utc)
    beijing_now = utc_now.astimezone(SHA_TZ)
    return str(beijing_now).split('.')[0].replace(' ','_').replace(':','-')