//
//  ViewController.m
//  XLSConvertToString
//
//  Created by anita on 2020/3/31.
//  Copyright © 2020 anita. All rights reserved.
//

#import "ViewController.h"
#import "DHxlsReader.h"

NSString *const excelPathStoreKey = @"excelPathStoreKey";
NSString *const stringPathStoreKey = @"stringPathStoreKey";
NSString *const projectResourceKey = @"projectResourceKey";
NSString *const langKeysKey = @"langKeysKey";

@interface ViewController()
@property (weak) IBOutlet NSTextField *excelPathField;
@property (weak) IBOutlet NSTextField *stringPathField;
@property (weak) IBOutlet NSTextField *projectStringDirPathField;
@property (weak) IBOutlet NSTextField *langKeysField;

@property (nonatomic, strong) NSArray *langKeys;
@end
@implementation ViewController

- (void)viewDidLoad {
    [super viewDidLoad];
    self.excelPathField.enabled = NO;
//    self.stringPathField.enabled = NO;
//    self.projectStringDirPathField.enabled = NO;
    [self.stringPathField setFocusRingType:NSFocusRingTypeNone];
    [self.projectStringDirPathField setFocusRingType:NSFocusRingTypeNone];
    [self.langKeysField setFocusRingType:NSFocusRingTypeNone];
    // Do any additional setup after loading the view.
    
    [self setupData];
    [self restoreBeforeConfigs];
}

#pragma mark - data
- (void)setupData{
    self.langKeys = @[@"en", @"es", @"de", @"fr", @"tr", @"it"];
    self.langKeysField.stringValue = [self.langKeys componentsJoinedByString:@","];
    
}

#pragma mark - restore
- (void)restoreBeforeConfigs{
    NSString *excelPathBefore = [[NSUserDefaults standardUserDefaults] objectForKey:excelPathStoreKey];
    if (excelPathBefore.length > 0){
        self.excelPathField.stringValue = excelPathBefore;
    }
    
    NSString *stringPathBefore = [[NSUserDefaults standardUserDefaults] objectForKey:stringPathStoreKey];
    if (stringPathBefore.length > 0){
        self.stringPathField.stringValue = stringPathBefore;
    }
    
    NSString *projectPathBefore = [[NSUserDefaults standardUserDefaults] objectForKey:projectResourceKey];
    if (projectPathBefore.length > 0){
        self.projectStringDirPathField.stringValue = projectPathBefore;
    }
    
    NSString *langKeysBefore = [[NSUserDefaults standardUserDefaults] objectForKey:langKeysKey];
    if (langKeysBefore.length > 0){
        self.langKeysField.stringValue = langKeysBefore;
        NSArray *langKeys = [langKeysBefore componentsSeparatedByString:@","];
        self.langKeys = langKeys;
    }
    
}

#pragma mark - action
- (IBAction)excelPathSelected:(NSButton *)sender {
    
    NSOpenPanel *panel = [NSOpenPanel openPanel];
    [panel setCanChooseFiles:YES];//是否能选择文件file
    [panel setCanChooseDirectories:YES];//是否能打开文件夹
    [panel setAllowsMultipleSelection:NO];//是否允许多选file
    [panel  setAllowedFileTypes:@[@"xls"]];
    NSInteger founded = [panel runModal]; //获取panel的响应
    if (founded == NSModalResponseOK) {
        NSArray * urls = [panel URLs];
        NSString * excelUrl = [[urls firstObject] path];
        self.excelPathField.stringValue = excelUrl;
        NSLog(@"excel path = %@",excelUrl);
        
        [[NSUserDefaults standardUserDefaults] setObject:excelUrl forKey:excelPathStoreKey];
        [[NSUserDefaults standardUserDefaults] synchronize];
    }
}

- (IBAction)stringFilePathSelected:(NSButton *)sender {
    
    NSOpenPanel *panel = [NSOpenPanel openPanel];
    [panel setCanChooseFiles:NO];//是否能选择文件file
    [panel setCanChooseDirectories:YES];//是否能打开文件夹
    [panel setAllowsMultipleSelection:NO];//是否允许多选file
    NSInteger founded = [panel runModal]; //获取panel的响应
    if (founded == NSModalResponseOK)
    {
        NSArray * urls = [panel URLs];
        NSString * stringUrl = [[urls firstObject] path];
        self.stringPathField.stringValue = stringUrl;
        NSLog(@"string path = %@",stringUrl);
        
        [[NSUserDefaults standardUserDefaults] setObject:stringUrl forKey:stringPathStoreKey];
        [[NSUserDefaults standardUserDefaults] synchronize];
    }
}

- (IBAction)projectStringDirPathSelect:(NSButton *)sender {
    NSOpenPanel *panel = [NSOpenPanel openPanel];
    [panel setCanChooseFiles:NO];//是否能选择文件file
    [panel setCanChooseDirectories:YES];//是否能打开文件夹
    [panel setAllowsMultipleSelection:NO];//是否允许多选file
    NSInteger founded = [panel runModal]; //获取panel的响应
    if (founded == NSModalResponseOK)
    {
        self.projectStringDirPathField.enabled = NO;
        NSArray * urls = [panel URLs];
        NSString * stringUrl = [[urls firstObject] path];
        self.projectStringDirPathField.stringValue = stringUrl;
        NSLog(@"project string path = %@",stringUrl);
        
        [[NSUserDefaults standardUserDefaults] setObject:stringUrl forKey:projectResourceKey];
        [[NSUserDefaults standardUserDefaults] synchronize];
    }
}


- (IBAction)convertButtonClicked:(NSButton *)sender {
    if(self.excelPathField.stringValue.length <= 0)
    {
        [self alert:@"请选择 xls 文件"];
        return;
    }
    if (self.stringPathField.stringValue.length <= 0 && self.projectStringDirPathField.stringValue.length <= 0) {
        [self alert:@"导出路径 和 项目翻译路径 至少选择一个"];
        return;
    }
    
    if (self.langKeysField.stringValue.length <= 0){
        [self alert:@"excel中对应语言的key需要指定（eg: en,es,de)"];
        return;
    }
    
    NSString *langKeysString = [self.langKeys componentsJoinedByString:@","];
    [[NSUserDefaults standardUserDefaults] setObject:langKeysString forKey:langKeysKey];
    [[NSUserDefaults standardUserDefaults] synchronize];
    
    NSLog(@"excel 路径 = %@",self.excelPathField.stringValue);
    DHxlsReader * reader = [DHxlsReader xlsReaderWithPath:self.excelPathField.stringValue];
    if (!reader)
    {
        [self alert:@"文件格式不正确!!!"];
        return ;
    }
    
    NSInteger sheetCount = [reader numberOfSheets];
    
    NSInteger row = [reader rowsForSheetAtIndex:0]; // 1983
    NSInteger cellCount = [reader numberOfColsInSheet:0]; // 16
    
    NSMutableArray * strings = [[NSMutableArray alloc] init];
    for (NSInteger i = 0; i < row; i++)
    {
        DHcell * keyCell = [reader cellInWorkSheetIndex:1 row:i + 1 col:1];
        if (keyCell.str.length == 0){ // key为空，跳过此翻译
            continue;
        }
        
        for (NSInteger j = 2; j <= cellCount; j++)
        {
            if (j > strings.count)
            {
                NSMutableString * str = [[NSMutableString alloc] init];
                [strings addObject:str];
            }
            NSMutableString * str = strings[j - 2];
            DHcell * cell = [reader cellInWorkSheetIndex:1 row:i + 1 col: j];
            if (cell.str.length > 0)
            {
                [str appendFormat:@"\n\"%@\" = ",keyCell.str];
                NSString * value = cell.str; // 可对取出来的值，进行一些额外处理。
                //value = [value stringByReplacingOccurrencesOfString:@"%s" withString:@"%@"];
                //value = [value stringByReplacingOccurrencesOfString:@"\"" withString:@"\\\""];
                if ([value containsString:@"\""]){
                    value = [value stringByReplacingOccurrencesOfString:@"\"" withString:@"\\\""];
                }
                value = [self removeSpaceAndNewLine:value];
                [str appendFormat:@"\"%@\";",value];
            }
            
        }
    }
    NSLog(@"sheetCount = %ld row = %ld cellCount = %ld",(long)sheetCount,row,cellCount);
    
    // ⚠️special deal for epal. 移除掉备注列
    if(strings.count > 0){
        NSMutableArray *tmp = [strings mutableCopy];
        [tmp removeObjectAtIndex:0];
        strings = tmp;
    }
    [self saveToFiles:strings];
}

- (void)saveToFiles:(NSArray *)strings
{

    
    for (NSInteger i = 0; i < strings.count; i++)
    {
        
        // strPath way 1 ：
//        NSString *langKey = self.langKeys[i];
//        // 创建文件文件夹。。放入文件
//        NSString *newDir = [self.stringPathField.stringValue stringByAppendingPathComponent:[NSString stringWithFormat:@"%@.lproj", langKey]];
//
//        if (![[NSFileManager defaultManager] fileExistsAtPath:newDir]){
//            NSError *createDirError;
//            [[NSFileManager defaultManager] createDirectoryAtPath:newDir withIntermediateDirectories:YES attributes:nil error:&createDirError];
//            if (createDirError){
//                [self alert:[NSString stringWithFormat:@"创建文件夹失败:%@ \n error = %@ ",newDir, createDirError.localizedDescription]];
//                return;
//            }
//        }
//        NSString * strPath = [newDir stringByAppendingPathComponent:@"Localizable.strings"];

        if (self.stringPathField.stringValue.length > 0){
            NSString * s = strings[i];
            if (s.length > 0)
            {
                NSString *stringName = self.langKeys[i];
                NSString * strPath = [self.stringPathField.stringValue stringByAppendingPathComponent:[NSString stringWithFormat:@"%@.strings", stringName]];
                NSData * data = [s dataUsingEncoding:NSUTF8StringEncoding];
                NSError * error ;
                BOOL writeResult = [data writeToURL:[NSURL fileURLWithPath:strPath] options:NSDataWritingAtomic error:&error];
                if (error){
                    NSLog(@"error = %@",[error userInfo]);
                }
                if (writeResult){
                    NSLog(@"写入文件成功");
                }
                else{
                    [self alert:[NSString stringWithFormat:@"写入文件失败:%@ \n error = %@",strPath,error.userInfo]];
                    NSLog(@"写入文件失败");
                    return ;
                }
            }
        }
        
        if (self.projectStringDirPathField.stringValue.length > 0){
            NSString * s = strings[i];
            if (s.length > 0)
            {
                NSString *langKey = self.langKeys[i];
                NSString *langDir = [self.projectStringDirPathField.stringValue stringByAppendingPathComponent:[NSString stringWithFormat:@"%@.lproj", langKey]];
                NSString *strPath = [langDir stringByAppendingPathComponent:@"Localizable.strings"];
                
                NSData * data = [s dataUsingEncoding:NSUTF8StringEncoding];
                NSError * error ;
                BOOL writeResult = [data writeToURL:[NSURL fileURLWithPath:strPath] options:NSDataWritingAtomic error:&error];
                if (error){
                    NSLog(@"error = %@",[error userInfo]);
                }
                if (writeResult){
                    NSLog(@"写入文件成功");
                }
                else{
                    [self alert:[NSString stringWithFormat:@"写入文件失败:%@ \n error = %@",strPath,error.userInfo]];
                    NSLog(@"写入文件失败");
                    return ;
                }
            }
        }
    }
    
    [self alert:@"操作完成"];
//    [[NSWorkspace sharedWorkspace] selectFile:nil inFileViewerRootedAtPath:self.stringPathField.stringValue];
}
- (void)setRepresentedObject:(id)representedObject {
    [super setRepresentedObject:representedObject];
    
    // Update the view, if already loaded.
}

#pragma mark - 不重要 action

- (IBAction)tapOpenExcelDir:(id)sender {
    if (self.excelPathField.stringValue.length > 0){
        [[NSWorkspace sharedWorkspace] selectFile:nil inFileViewerRootedAtPath:self.excelPathField.stringValue];
    }
}
- (IBAction)tapOpenStringDir:(id)sender {
    if (self.stringPathField.stringValue.length > 0){
        [[NSWorkspace sharedWorkspace] selectFile:nil inFileViewerRootedAtPath:self.stringPathField.stringValue];
    }
}

- (IBAction)tapOpenProjectStringDir:(id)sender {
    if (self.projectStringDirPathField.stringValue.length > 0){
        [[NSWorkspace sharedWorkspace] selectFile:nil inFileViewerRootedAtPath:self.projectStringDirPathField.stringValue];
    }
}


#pragma mark - helper
- (void)alert:(NSString *)msg{
    NSAlert *alert = [[NSAlert alloc] init];
    alert.messageText = @"系统提示:";
    alert.informativeText = msg;
    [alert addButtonWithTitle:@"确定"];
    NSInteger ret = [alert runModal];
    switch(ret){
        default:
            printf("点击了OK。\n");
            break;
    }
}

- (NSString *)removeSpaceAndNewLine:(NSString *)str {
    NSString *temp = [str.copy stringByTrimmingCharactersInSet:[NSCharacterSet whitespaceCharacterSet]];
    NSString *text = [temp stringByTrimmingCharactersInSet:[NSCharacterSet whitespaceAndNewlineCharacterSet ]];
    return text;
}


@end
