#coding=utf-8
'''
  检查PC的领用、借用、归还、入库等事件记录
@author: junli
'''


from collections import defaultdict
from operator import  attrgetter
from openpyxl.reader.excel import load_workbook



class EventRecordRow:
    
    headers = "日期    变更类型    单据编号    品牌    型号    序列号    快速服务代码    使用部门    使用人    备注".split()
    
    @classmethod
    def header_length(cls):
        return len(cls.headers)
    
    def __init__(self,values):
        self.values = values
        
    @property
    def event_date(self):
        return self.values[0]

    @property
    def seq_no(self):
        return self.values[5]
    
    @property
    def sheet_no(self):
        return self.values[2]
    
    @property
    def event_type(self):
        return self.values[1]


class EventsValidator:
    NEXT_EVENTS = {
            '借用' : ['归还', '入库'],
            '领用' : [ '入库'],
            '归还': ['借用', '领用'],
            '入库': ['借用', '领用']
        }
    
    @staticmethod
    def is_events_valide(event_types):
        if len(event_types) > 1:
            for i in range(len(event_types)-1):
                next_event = event_types[i+1]
                if next_event not in EventsValidator.NEXT_EVENTS[event_types[i]]:
                    return False
        return True


class PCEventsStore:
    
    def __init__(self):
        self.seq_events = {}
        self.sheet_events = {}
        self.incomplete_events = []
        self.count = 0
        
    
    def load_from_file(self,filepath):
        wb = load_workbook(filepath)
        ws = wb['变更记录']
        
        sheet_events = defaultdict(list)
        seq_events = defaultdict(list)
        incomplete_events = []
        record_col_count = EventRecordRow.header_length()
        count = 0
        for row in ws.rows:           
            # skip header row
            if row[0].value=='日期':
                continue                 
            # empty row: stop
            if row[0].value is None and row[1].value is None and row[2].value is None:
                break            
            
            values = [cell.value for cell in row[:record_col_count]]
            record =  EventRecordRow(values)
            # if no seq_no, bad rows 
            if record.seq_no is None or record.seq_no=='未带待核实':
                incomplete_events.append(record)
            else:
                seq_events[record.seq_no].append(record)
            
            if record.sheet_no is not None:
                sheet_events[record.sheet_no].append(record)            
            count +=1                    
                
        self.sheet_events = sheet_events
        self.seq_events = seq_events
        self.incomplete_events = incomplete_events
        self.count = count


    
    def validate_events_sequences(self):       
        bad_pattern_seqs = defaultdict(list)
        lack_sheet_seqs = []
        for seq,events in self.seq_events.items():
            # sort by date
            events.sort(key=attrgetter('event_date'))
            event_types = [event.event_type for event in events]
            if not EventsValidator.is_events_valide(event_types):
                pattern = "-".join(event_types)
                bad_pattern_seqs[pattern].append(seq)
            
            last_event = events[-1]
            if event_types[-1] in ['借用', '领用'] and last_event.sheet_no is None:
                lack_sheet_seqs.append(last_event)
                
                
        return bad_pattern_seqs,lack_sheet_seqs
       
def main():
    filepath = r'testdata\pcs.xlsx'
    store = PCEventsStore()
    store.load_from_file(filepath)
    print("Load {} events from file {}; Bad records: {}".format(store.count, filepath, len(store.incomplete_events)))
    print("... about {} machines.".format(len(store.seq_events)))
    
    bad_pattern_seqs,lack_sheet_seqs = store.validate_events_sequences()
    print("Bad sequence of events[{}]:".format(len(bad_pattern_seqs)))
    for p,seqs in bad_pattern_seqs.items():
        print("{} : {}".format(p, ','.join(seqs)))
        
    print("Those machines are lacking a sheet no:")
    for event in lack_sheet_seqs:
        print(event.values)

    

if __name__ == '__main__':    
    main()