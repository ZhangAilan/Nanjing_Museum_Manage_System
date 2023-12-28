import folium

#创建一个Folium地图对象
m = folium.Map(location=[32.04,118.78], zoom_start=13, tiles='CartoDB Positron')

#博物馆经纬度
museums={
    "南京博物院":[32.04446888, 118.8209673],
    "南京市博物总馆":[32.03639749, 118.7701077],
    "中共代表团梅园新村纪念馆":[32.04463692,118.7965663],
    "六朝博物馆":[32.045,118.79379],
    "中国科举博物馆":[32.02334,118.78549],
    "渡江胜利纪念馆":[32.07598,118.72650],
    "南京民俗博物馆":[32.02793,118.77659],
    "江宁织造博物馆":[32.04457,118.78898],
    "太平天国历史博物馆":[32.02311,118.78004],
    "南京地质博物馆":[32.04673,118.80205],
    "南京古生物博物馆":[32.06164,118.79011],
    "南京云锦博物馆":[32.03823,118.73963],
    "南京城墙博物馆":[32.01431,118.77766]
}

#在地图上添加地标
for museum,location in museums.items():
    folium.Marker(location,
                  popup=museum,
                  icon=folium.Icon(color='red',icon='info-sign')
                  ).add_to(m)

#将地图保存为html文件
map_file='NanjingCityMap.html'
m.save(map_file)