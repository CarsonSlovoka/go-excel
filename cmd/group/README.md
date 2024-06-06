如果你的資料類型是

```
type Group struct {
	Name  string
	Items []Item
}

type Item struct {
	Ch      rune
	Unicode string
}
```

其中Item的欄位可以自訂(沒有slice, map, struct)

那麼此範例可以很好的視覺化`[]Group`的資料內容
