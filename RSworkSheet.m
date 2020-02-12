//
//  workSheet.m
//  excelWriterExample
//
//  Created by andrea cappellotto on 14/09/11.
//  Copyright 2011 Università degli studi di Trento. All rights reserved.
//

#import "RSworkSheet.h"

@implementation RSworkSheet
@synthesize columnWidth, rowHeight,name, arrayWorkSheetRow;


- (id)init
{
    self = [super init];
    if (self) 
    {
        // Initialization code here.
    }
    
    return self;
}

- (void)addWorkSheetRow:(RSworkSheetRow*)row
{
    [arrayWorkSheetRow addObject:row];
}
- (id)initWithName:(NSString*)nameSheet
{
    self = [super init];
    if (self) 
    {
        name = [[NSString alloc] initWithFormat:@"%@",nameSheet];
        columnWidth = 85;
        rowHeight=20;
        arrayWorkSheetRow = [[NSMutableArray alloc] init];
    }
    
    return self; 
}


-(void)dealloc
{
    
}
@end
