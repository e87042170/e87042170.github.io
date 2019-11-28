<h2>最新文章</h2>
<ul>
  {% for post in site.posts %}
    <li>
      <h2><a href="{{ post.url }}">{{ post.title }}</a></h2>
      <p><small>{{ post.date | date_to_string }}</small></p>
      <p>{{ post.excerpt }}</p>
    </li>
  {% endfor %}
</ul>
