const assert=require('assert');
const utils=require('../src/utils');



describe('test uitls',function(){


    describe('test #charToNum()',function(){

        it('single char : [A-Z]|[a-z]',function(){
            const tests={
                'A': 1,
                'Z': 26,
                'a': 97-65+1,    
                'z': 97-65+26,
            };
            Object.keys(tests).forEach(key=>{
                const v=tests[key];
                assert.equal(utils.charToNum(key),v,`code ${key} to num does not equal the expected value ${v}`);
            });
        });

        it('other case ( :todo )',function(){
            // todo:
        });

    });

    describe('test #charPlus()',function(){
        
        it('single char : [A-Z]|[a-z]',function(){
            const tests={
                'A': 'B',
                'Z': 'AA',
                // 'a': '',    don't make sense . might be modified in the future
                // 'z': 97-65+26,
            };
            const delta=1;
            Object.keys(tests).forEach(key=>{
                const expected=tests[key];
                const actual=utils.charPlus(key,delta);
                assert.equal(actual,expected,`char ${key} + delta ${delta} doesn't make sense`);
            });
        });

        it('other case ( :todo )',function(){
        });

    });
});