const fs = require("fs");
const path=require('path');
const assert=require('assert');
const {_asyncToGenerator,Promise_fromCallback,Promise_fromStandard}=require('../src/async');


describe('test async.js',function(){

    describe('test #Promise_fromCallback',function(){

        const data='hello,world';

        function fn(cb){
            setTimeout(() => {
                cb(data);
            }, 1000);
        }

        it('simple test ',function(){

            const f=Promise_fromCallback(fn);
            return f()
                .then(function(str){
                    assert.ok(str,data,'wrong parameter received');
                });
        });
    });

    describe('test #Promise_fromStand',function(){
        
        it('should reject',function(){
            const err=new Error('shit happens');
            const data='hello,world';
            function fn(cb){
                setTimeout(() => {
                    cb(err,data);
                }, 1000);
            }
            const f=Promise_fromStandard(fn);
            return f().then(
                data=>{assert.fail(`callback(err,data) should reject when err given `);},
                e=>{ 
                    assert.equal(e.message,err.message,'should throw the same error message');
                    assert.ok('callback(err,data) should reject');
                }
            );
        });
        it('should resolve',function(){
            const err=null;
            const data='hello,world';
            function fn(cb){
                setTimeout(() => {
                    cb(err,data);
                }, 1000);
            }
            const f=Promise_fromStandard(fn);
            return f().then(
                d=>{
                    assert.equal(d,data,'should receive the same data');
                    assert.ok('callback(null,data) should resolve');
                },
                e=>{ 
                    assert.fail(`callback(err,data) should resolve when err is null/undefined .etc`);
                }
            );
        });
    });

    describe('test #_asyncToGenerator',function(){
        
        it('simple test',function(){
            let i=0;
            function* fn(){
                yield ++i;
                yield ++i;
                yield ++i;
                yield ++i;
            }
            const f=_asyncToGenerator(fn);
            return f().then(final=>{
                assert.equal(i,4,"i should equals 4 when final yield executed");
            });
        });
    });
});