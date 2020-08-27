import hmac
import hashlib
import base64

def string_to_hash(word):
    word = word.encode('utf-8')
    hash = hmac.new(word, word, hashlib.sha1).digest()
    return base64.b64encode(hash).decode("utf-8")

print(string_to_hash('a')) #OQLthH/yiTC18UGr+otHFoElNnM=