/**
 * 渲染excel模板
 * @param exlBuf 模板excel的Buff
 * @param _data_ 数据
 */
export declare function renderExcel(exlBuf: Buffer, _data_: any): Promise<Buffer>;

/**
 * 渲染excel模板
 * @param exlBuf 模板excel的Buff
 * @param _data_ 数据
 * @param cb 渲染成功后回调
 */
export declare function renderExcelCb(exlBuf: Buffer, _data_: any, cb: Function): void;

/**
 * 渲染excel模板
 * @param exlBuf 模板excel的Buff
 * @param _data_ 数据
 * @param opt 选项, cachePath 编译缓存路径
 */
export declare function renderExcel(exlBuf: Buffer, _data_: any, opt: { cachePath?: string }): Promise<Buffer>;

/**
 * 渲染excel模板
 * @param exlBuf 模板excel的Buff
 * @param _data_ 数据
 * @param cb 渲染成功后回调
 */
export declare function renderExcelCb(exlBuf: Buffer, _data_: any, opt: { cachePath?: string }, cb: Function): void;
