import React, { useState } from 'react'

const FieldMapping = () => {
    const [relation,setRelation]= useState([{source:{},target:{}}]);
    const [sourceData,setSourceData] = useState(new Array(7).fill().map((item, idx) => ({
        name: `field${idx + 1}`,
        type: 'string',
        key: `field${idx + 1}`,
        desc: `这是表字段field${idx + 1}`,
        operate: `测试${idx}`
      })));
    const sourceCols = [
        { title: 'Field Name', key: 'name', width: '80px' },
        { title: 'Action', key: 'operate', width: '80px', align: 'center', render: (value, record) => {
          return <a href="javascript:void(0);" onClick={
            () => {
              alert(JSON.stringify(record));
            }
          }>Alert</a>;
        }}
      ];
      const targetCols = [
        { title: 'Field Name', key: 'name', width: '50%' },
        { title: 'Action', key: 'type', width: '50%' }
      ];
    return (
        <div>
            
        </div>
    )
}

export default FieldMapping
