//
//  NSStyle.m
//  excelWriterExample
//
//  Created by andrea cappellotto on 14/09/11.
//  Copyright 2011 Università degli studi di Trento. All rights reserved.
//

#import "RSStyle.h"

@implementation RSStyle
@synthesize font, size, alignmentH, alignmentV, color;
- (id)init
{
    self = [super init];
    if (self) {
        // Initialization code here.
    }
    
    return self;
}
-(void)dealloc
{
    [super dealloc];
}

@end
