## Purpose
Make SharePoint 2013/2016/Online list item edits easy from JavaScript with a common library that generates HTTP traffic formatted with correct headers and verbs. Enjoy!


## Get Started (AngularJS / spcrud.js)
1. upload files to /SiteAssets/
2. add Content Editor to homepage
3. point Content Editor at /SiteAssets/spcrud-demo.html
4. create new Custom List (ex: "Test")
5. click buttons to run CRUD operations on list "Test" (or any list)

## Get Started (Angular2+ / spcrud.ts)
1. Import module with 

```
import { Spcrud } from './spcrud';
```
2. Add provider with
```
providers: [Spcrud]
```

3. Add contstructor with 
```
constructor(private spcrud: Spcrud) {
    this.spcrud.setBaseUrl('http://portal/sites/todo');
}
```

4. Add save method with
```
saveList() {
    const ctl = this;
    this.spcrud.jsonWrite('jsonTodo', this.todos).then(function (r) {
        ctl.status = r;
    });
}
```

## Video
[![](https://raw.githubusercontent.com/spjeff/spcrud/master/doc/7.png)](http://www.spjeff.com/2016/04/10/video-get-started-with-spcrud-js/ "Get Started with SPCRUD")

## Sceenshots
1. upload files to /SiteAssets/
![image](https://raw.githubusercontent.com/spjeff/spcrud/master/doc/1.png)
2. homepage with Content Editor pointed at /SiteAssets/spcrud-demo.html
![image](https://raw.githubusercontent.com/spjeff/spcrud/master/doc/2.png)
3. Create
![image](https://raw.githubusercontent.com/spjeff/spcrud/master/doc/3.png)
4. Read
![image](https://raw.githubusercontent.com/spjeff/spcrud/master/doc/4.png)
5. Update
![image](https://raw.githubusercontent.com/spjeff/spcrud/master/doc/5.png)
6. Delete
![image](https://raw.githubusercontent.com/spjeff/spcrud/master/doc/6.png)

## Contact
Please drop a line to [@spjeff](https://twitter.com/spjeff) or [spjeff@spjeff.com](mailto:spjeff@spjeff.com)
Thanks!  =)

![image](http://img.shields.io/badge/first--timers--only-friendly-blue.svg?style=flat-square)


## License

The MIT License (MIT)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.