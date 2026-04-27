import base64
import getpass
import hashlib
import secrets


ITERATIONS = 260_000


def hash_password(password):
    salt = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("utf-8"),
        ITERATIONS,
    )
    encoded = base64.b64encode(digest).decode("ascii")
    return f"pbkdf2_sha256${ITERATIONS}${salt}${encoded}"


if __name__ == "__main__":
    password = getpass.getpass("Password da autorizzare: ")
    confirm = getpass.getpass("Ripeti password: ")
    if password != confirm:
        raise SystemExit("Le password non coincidono.")
    if len(password) < 8:
        raise SystemExit("Usa almeno 8 caratteri.")
    print(hash_password(password))
