const soap = require('soap')
const XLSX = require('xlsx')
const fs = require('fs')

const url = 'http://s70-1c04.trk.tom.ru/Services/ws/PhoneList.1cws?wsdl'
const args = {}
const parseString = require('xml2js').parseString

// const creds = new soap.BasicAuthSecurity('username', 'password');

const exportToXLS = (data, filename = 'out.xlsb') => {
  const workbook = XLSX.utils.book_new()
  // const worksheet = XLSX.utils.aoa_to_sheet([[]])
  // XLSX.utils.sheet_add_aoa(worksheet, [
  //   [1], '',
  //   ['1111', 2222, 3333]
  // ], { origin: 'B3' })

  const worksheet = XLSX.utils.json_to_sheet(data)
  // XLSX.utils.json_to_sheet(worksheet, [{ a: 1 }], { origin: 'B2' })
  XLSX.utils.book_append_sheet(workbook, worksheet, 'sheet_name')
  XLSX.writeFile(workbook, filename)
}

function listToTree (list) {
  const map = {}; let node; const roots = []; let i

  for (i = 0; i < list.length; i += 1) {
    map[list[i].id] = i // initialize the map
    list[i].children = [] // initialize the children
  }

  for (i = 0; i < list.length; i += 1) {
    node = list[i]
    if (node.parent_id !== '') {
      // if you have dangling branches check that map[node.parentId] exists
      if (!list[map[node.parent_id]]) {
        // console.log(list)
        // console.log(node)
      } else { list[map[node.parent_id]].children.push(node) }
    } else {
      roots.push(node)
    }
  }
  return roots
}

const treeToList = (tree) => {
  const result = []
  if (tree) {
    tree.forEach((item) => {
      result.push(item)
      treeToList(item.children).forEach((child) => child.Name && result.push([item.Name + ' - ' + child.Name]))
    })
  }
  return result
}

const fillNames = (item, ar) => {
  // console.log(item, ar)
  if (item.parent_id && ar[item.parent_id]) {
    return fillNames(ar[item.parent_id], ar) + ' - ' + item.Name
  }
  return item.Name
}

const fetchDivision = () => {
  soap.createClient(url, { wsdl_options: { auth: { username: 'robot01', password: 'robot01' } } }, function (err, client) {
    if (err) console.log(err)
    client.setSecurity(new soap.BasicAuthSecurity('robot01', 'robot01'))
    client.GetPodrs(args, function (err, response) {
      if (err) console.log(err)
      const content = Buffer.from(response.return, 'base64').toString('utf-8')
      fs.writeFile('dataPodrs.xml', content, () => {
        console.log()
      })
      parseString(content, (err, result) => {
        if (err) console.log(err)
        const ao = []
        // result.Podrs.Podr.forEach((i) => {
        //   ao[i.ID[0]] = { id: i.ID[0], Name: i.Name[0], parent_id: i.ParentID[0], children: [] }
        // })
        result.Podrs.Podr.forEach((i) => ao.push({ id: i.ID[0], Name: i.Name[0], parent_id: i.ParentID[0], children: [] }))
        const aa1 = ao.reduce((all, i) => {
          all[i.id] = i
          return all
        }, {})

        const ddo = ao.map((item) => ({ ...item, FullName: fillNames(item, aa1) }))

        // const bo = listToTree(ao)
        // const ao = result.Podrs.Podr
        // console.log(bo, typeof bo)
        // const co = treeToList(bo)
        // console.log(aa1)
        exportToXLS(ddo, 'divisons.xls')
        // result.Podrs.Podr.map((item) => {
        //   console.log(item)
        //   return 0
        // })
      })
    })
  })
}

const fetchUsers = () => {
  soap.createClient(url, { wsdl_options: { auth: { username: 'robot01', password: 'robot01' } } }, function (err, client) {
    if (err) console.log(err)
    client.setSecurity(new soap.BasicAuthSecurity('robot01', 'robot01'))
    client.GetSotrs(args, function (err, response) {
      if (err) console.log(err)
      const content = Buffer.from(response.return, 'base64').toString('utf-8')
      fs.writeFile('dataSotrs.xml', content, () => {
        console.log()
      })
      parseString(content, (err, result) => {
        if (err) console.log(err)
        const ao = []
        // result.Podrs.Podr.forEach((i) => {
        //   ao[i.ID[0]] = { id: i.ID[0], Name: i.Name[0], parent_id: i.ParentID[0], children: [] }
        // })
        result.Sotrs.Sotr.forEach((i) => ao.push({
          id: i.ID[0],
          // guid: i.GUID[0],
          name: i.Name[0],
          podrID: i.PodrID[0],
          PodrName: i.PodrName[0],
          Dol: i.Dol[0],
          Sost: i.Sost[0],
          Date1: i.Date1[0],
          Date2: i.Date2[0],
          Email: i.Email[0],
          Addr: i.Addr[0],
          WorkDate: i.WorkDate[0]
        }))

        // const ddo = ao.map((item) => ({ ...item, FullName: fillNames(item, aa1) }))

        // const bo = listToTree(ao)
        // const ao = result.Podrs.Podr
        // console.log(bo, typeof bo)
        // const co = treeToList(bo)
        // console.log(aa1)
        exportToXLS(ao, 'users.xls')
        // result.Podrs.Podr.map((item) => {
        //   console.log(item)
        //   return 0
        // })
      })
    })
  })
}

fetchDivision()
fetchUsers()
