#
# Created by Berke Akyıldız on 05/July/2019
#
import msoffcrypto

file = msoffcrypto.OfficeFile(open("C:\\Users\\MONSTER\\Desktop\\asd.docx", "rb"))

# Use password
file.load_key(password="Passw0rd")

# Use private key
# file.load_key(private_key=open("priv.pem", "rb"))
# Use intermediate key (secretKey)
# file.load_key(secret_key=binascii.unhexlify("AE8C36E68B4BB9EA46E5544A5FDB6693875B2FDE1507CBC65C8BCF99E25C2562"))

file.decrypt(open("C:\\Users\\MONSTER\\Desktop\\decrypted.docx", "wb"))
