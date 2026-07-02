import paramiko


def connect_ontap(ip, username, password):

    ssh = paramiko.SSHClient()

    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    ssh.connect(
        hostname=ip,
        username=username,
        password=password
    )

    return ssh
