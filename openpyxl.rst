Openpyxl
========
Opening Excel Documents with OpenPyXL
-------------------------------------
>>> import openpyxl
>>> wb = openpyxl.load_workbook('example.xlsx')  #get a workbook

link to shutil_ “Why shouldn’t I work for the NSA? That’s a tough one, but I’ll take a shot. Say I’m working at the NSA and somebody puts a code on my desk, something no one else can break. Maybe I take a shot at it, and maybe I break it. I’m real happy with myself, ‘cause I did my job well. But maybe that code was the location of some rebel army in North Africa or the Middle East and once they have that location they bomb the village where the rebels are hiding. Fifteen hundred people that I never met, never had no problem with, get killed.

 

Now the politicians are saying ‘Oh, send in the Marines to secure the area,’ ‘cause they don’t give a shit. It won’t be their kid over there getting shot just like it wasn’t them when their number got called ‘cause they were pulling a tour in the National Guard. It’ll be some kid from Southie over there taking shrapnel in the ass. He comes back to find that the plant he used to work at got exported to the country he just got back from, and the guy that put the shrapnel in his ass got his old job, ‘cause he’ll work for fifteen cents a day and no bathroom breaks.

.. _shutil:

Meanwhile he realizes that the only reason he was over there in the first place was so we could install a government that would sell us oil at a good price. And of course the oil companies use the little skirmish to scare up domestic oil prices, a cute little ancillary benefit for them, but it ain’t helping my buddy at two-fifty a gallon. They’re taking their sweet time bringing the oil back, of course, and maybe they took the liberty of hiring an alcoholic skipper who likes to drink martinis and fucking play slalom with the icebergs. It ain’t too long until he hits one, spills the oil, and kills all the sea life in the North Atlantic.

Public Key Cryptography
All of the ciphers in this book so far have one thing in common: the key used for encryption is the same key used for decryption. This leads to a tricky problem: How do you share encrypted messages with someone you’ve never talked to before?

Say someone on the other side of the world wants to communicate with you. But you both know that spy agencies are monitoring all emails, letters, texts, and calls that you send. You could send them encrypted messages, however you would both have to agree on a secret key to use. But if one of you emailed the other a secret key to use, then the spy agency would be able to see this key and then decrypt any future messages you send with that key. Normally you would both secretly meet in person and exchange the key then. But you can’t do this if the person is on the other side of the world. You could try encrypting the key before you send it, but then you would have to send the secret key for that message to the other person and it would also be intercepted.

This is a problem solved by public key cryptography. Public key cryptography ciphers have two keys, one used for encryption and one used for decryption. A cipher that uses different keys for encryption and decryption is called an asymmetric cipher, while the ciphers that use the same key for encryption and decryption (like all the previous ciphers in this book) are called symmetric ciphers.

The important thing to know is that a message encrypted with one key can only be decrypted with the other key. So even if someone got their hands on the encryption key, they would not be able to read an encrypted message because the encryption key can only encrypt; it cannot be used to decrypt messages that it encrypted.

So when we have these two keys, we call one the public key and one the private key. The public key is shared with the entire world. However, the private key must be kept secret.

If Alice wants to send Bob a message, Alice finds Bob’s public key (or Bob can give it to her). Then Alice encrypts her message to Bob with Bob’s public key. Since the public key cannot decrypt a message that was encrypted with it, it doesn’t matter that everyone else has Bob’s public key.

When Bob receives the encrypted message, he uses his private key to decrypt it. If Bob wants to reply to Alice, he finds her public key and encrypts his reply with it. Since only Alice knows her own private key, Alice will be the only person who can decrypt the encrypted message.




this is an example.
 
