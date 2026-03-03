import hmac
import hashlib
import os

BASE_URL = "https://equatorportal-calculator.onrender.com"

CUSTOMER_LINK_SECRET = os.getenv("CUSTOMER_LINK_SECRET")

def generate_signature(customer_id: str):
    msg = customer_id.encode("utf-8")
    secret = CUSTOMER_LINK_SECRET.encode("utf-8")

    sig = hmac.new(secret, msg, hashlib.sha256).hexdigest()
    return sig


def generate_link(customer_id: str):
    sig = generate_signature(customer_id)

    url = f"{BASE_URL}/?c={customer_id}&sig={sig}"

    print("\nCustomer ID:", customer_id)
    print("Signature :", sig)
    print("Link      :", url)


if __name__ == "__main__":
    customer = input("Customer ID: ").strip()
    generate_link(customer)
