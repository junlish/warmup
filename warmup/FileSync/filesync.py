'''
Created on 2018年6月29日
    Compare source dir and dest dir, report the difference.
        If --sync is given, then copy diff files from source to dest, and remove obsolete files from dest.
@author: junli
'''

import filecmp
import os.path
import argparse
import shutil

from collections import namedtuple,Counter

DiffItem = namedtuple("DiffItem", ["dir1","dir2","item","type","is_dir","size"])
#types
MISSING = "Missing"
OBSOLETE = "Obsolete"
CHANGED = "Changed"
COMPARE_TYPES=[MISSING, CHANGED,OBSOLETE]

IGNORE_LIST= list(filecmp.DEFAULT_IGNORES) + ['.picasa']



def compare_main_dest(main,dest):
    diffs = []
    dcmp = filecmp.dircmp(main,dest,IGNORE_LIST)
    # missing items in dest dir
    for item in dcmp.left_only:
        file = os.path.abspath(os.path.join(main, item))
        file_d = os.path.abspath(os.path.join(dest, item))  
        is_dir = os.path.isdir(file)
        total_size = 0
        if is_dir:           
            for dirpath, _, filenames in os.walk(file):
                for f in filenames:
                    fp = os.path.join(dirpath, f)
                    total_size += os.path.getsize(fp)
        else:
            total_size = os.path.getsize(file)
        diffs.append(DiffItem(file,file_d,item,MISSING,is_dir,total_size))           
    
    # obsolete items in dest dir
    for item in dcmp.right_only:
        dest_file =  os.path.abspath(os.path.join(dest,item))
        diffs.append(DiffItem(main,dest_file,item,OBSOLETE, os.path.isdir(dest_file) ,0))
    
    # files changed    
    for item in dcmp.diff_files:
        file = os.path.abspath(os.path.join(main, item))
        file_d = os.path.abspath(os.path.join(dest, item))  
        diffs.append(DiffItem(file,file_d,item,CHANGED,False,os.path.getsize(file) ))
        
    # sub directories: recursively compare each sub dirs
    for comm_dir in dcmp.common_dirs:
        main_sub = os.path.abspath(os.path.join(main, comm_dir))
        dest_sub = os.path.abspath(os.path.join(dest, comm_dir))
        diffs.extend(compare_main_dest(main_sub, dest_sub))
        
    return diffs


def sync_files(diff_items, remove_obsolete):
    print("Synchronizing {} items...".format(len(diff_items)))
    
    total = len(diff_items)
    for i,item in enumerate(diff_items):
        if item.type==MISSING:
            print("[{}/{}] Copying :  {} --> {}".format(str(i+1),str(total), item.dir1,item.dir2))
            if item.is_dir:
                shutil.copytree(item.dir1, item.dir2)
            else:
                shutil.copy2(item.dir1,item.dir2)
        
        elif item.type == CHANGED:
            print(" [{}/{}]: Overwriting:  {} --> {}".format(str(i+1),str(total), item.dir1,item.dir2))
            shutil.copy2(item.dir1,item.dir2)
        elif item.type == OBSOLETE:
            if not remove_obsolete:
                print("[{}/{}]: Skipping obsolete file {}".format(str(i+1),str(total),item.dir2))
            else:
                print("[{}/{}]: Removing file {}".format(str(i+1),str(total),item.dir2))
                if item.is_dir:
                    shutil.rmtree(item.dir2)    
                else:
                    os.remove(item.dir2)
        else:
            raise Exception("Unknow action type"+str(item.type))
    
    

def main():
    parser = argparse.ArgumentParser(description='Compare & sync two folders')
    parser.add_argument('--source', required=True, help='Source Directory')
    parser.add_argument('--dest', required=True, help='Destination Directory')
    parser.add_argument('-s','--sync',action='store_true', help="Sychronize files from source to dest")
    parser.add_argument('-o','--obsolete',action='store_true', help="Remove obsolete files from dest dir" )
    parser.add_argument('-v','--verbose',action='store_true', help="Print more information" )
    args = parser.parse_args()
    
    #print(args)
    #diffs = compare_main_dest(r"D:\pictures\coloring", r"C:\temp\coloring")
    diffs = compare_main_dest(args.source, args.dest)
    
    if len(diffs)==0:
        print("No difference.")
        return
    
    counter = Counter()
    total_size =0
    for d in diffs:
        counter[d.type] +=1
        total_size += d.size
         
    compare_ret_str = ", ".join(["{} files:{}".format(t,counter[t]) for t in COMPARE_TYPES ]   )
    print(compare_ret_str)    
    print("Total file size: {:.2f}M".format(total_size/1024/1024))
    
    if args.verbose:
        print("Details:")
        for d in diffs:               
            print(d)
       
    
    if args.sync:
        sync_files(diffs, args.obsolete)

if __name__ == '__main__':
    main()
    
    
  
    
    