# MSG Reader

Javascript Module version of https://github.com/ykarpovich/msg.reader
MSG Reader is an Outlook Item File (.msg) reader that is built with HTML5.
Allows to parse and extract necessary information (attachment included) from .msg file.

Check [netas-ch/eml-parser](https://github.com/netas-ch/eml-parser) if you want to read .eml files.


# demo
check https://raw.githack.com/netas-ch/msg.reader/main/_test/test.html for a demo.


# usage
    const msgr = await import('./src/MsgReader.js');
    const email = new msgr.MsgReader(fileAsArrayBuffer);

    console.log(email.getDate());
    console.log(email.getSubject());
    console.log(email.getFrom());
    console.log(email.getCc());
    console.log(email.getTo());
    console.log(email.getReplyTo());
    console.log(email.getAttachments());

# license
Forked from [ykarpovich/msg.reader](https://github.com/ykarpovich/msg.reader) ; Copyright 2021 Yury Karpovich
Modified by Lukas Buchs, netas.ch

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.