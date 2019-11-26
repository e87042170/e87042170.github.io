## HTML 工具

[Cable Tray 容量計算](https://e87042170.github.io/CableTrayCalculator/) 

## HTML 遊戲

[2D Breakout Game](https://e87042170.github.io/2DBreakoutGame/) 

[Hanoi Tower 遞迴 - 河內塔](https://e87042170.github.io/HanoiTower/) 

## Posts

<ul>
  {% for post in site.posts %}
    <li>
      <a href="{{ post.url }}">{{ post.title }}</a>
    </li>
  {% endfor %}
</ul>

{% for post in site.posts %}
 <section class="section">
  <article>
    <div class="page-header">
      <h1><a href="{{ BASE_PATH }}{{ post.url }}">{{ post.title }}</a><h1>
    </div>
    <div class="note post-info">
      分類：<a href="categorise.html#{{ post.category }}-ref">{{ post.category}}</a>
    </div>
        
    {% if post.content contains "<!-- more -->" %}
      {{ post.content | split:"<!-- more -->" | first % }}
    {% else %}
      {{ post.content | strip_html | truncatewords:100 }}
    {% endif %}
    
    <div class="read-more">
      <a class="btn" href="{{ BASE_PATH }}{{ post.url }}">Read more...</a>
    </div>
  <article>
</section>
{% endfor %}

{% include adsense.html %}
