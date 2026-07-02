def execute_command(ssh, command):

    stdin, stdout, stderr = ssh.exec_command(command)

    output = stdout.read().decode()

    error = stderr.read().decode()

    if error:
        return error

    return output
