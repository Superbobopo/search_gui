"""Check the two screenshot identities using the Desktop cache read-only."""
from pathlib import Path
import sys
import threading

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from occurrence_comparison import AnalysisCache
from tools.check_real_case import ReadOnlyStore
from trace_sources import desktop_directory, robot_scope
from trace_store import cache_path


def main():
    path=cache_path([str(Path(desktop_directory())/'traces')])
    if not path.is_file():
        raise SystemExit('Cache Bureau existant requis ; aucune indexation ne sera lancée.')
    before=path.stat()
    # No constructor, schema installation, history initialization or synchronization.
    store=ReadOnlyStore.__new__(ReadOnlyStore)
    store.database=path
    store._local=threading.local()
    store._analysis_cache=AnalysisCache()
    cases=(('R1','1054','A0118427AA','292226','260210.004','13|13:40:44'),
           ('R2','1052','A0110543AC','343132','261012.095','13|13:37:15'))
    for robot,ident,reference,sn,lot,stamp in cases:
        with store.snapshot():
            rows=store._rows('''SELECT l.*,f.path FROM lines l JOIN files f ON f.id=l.file_id
                WHERE l.modbus=? AND l.ref=? AND l.sn=? AND l.lot=?''',(ident,reference,sn,lot))
            matches=[r for r in rows if robot_scope(r['path'])[1]==robot and stamp in r['text']]
            assert matches, (robot,ident,'absent du cache')
            for row in matches:
                result=store.comparison_for_occurrence(row['path'],row['number'])
                assert result['different']==(ident=='1054'), (robot,ident,result)
                if ident=='1054':
                    assert result['pair'] and robot_scope(result['pair'][1]['path'])[1]=='R2'
                else:
                    assert result['status'] in ('missing_measurement','ambiguous_measurement')
                print(robot,ident,result['status'],'red=',result['different'])
    after=path.stat()
    assert (before.st_size,before.st_mtime_ns)==(after.st_size,after.st_mtime_ns)
    print('OK cache reel en lecture seule, taille/date inchangees, aucune source relue ou modifiee')


if __name__=='__main__':
    main()
