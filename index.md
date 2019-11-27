<h2>Posts:</h2>
<ul>
  {% for post in site.posts %}
    <li>
      <small>{{ post.date | date_to_string }}</small>
      <h2><a href="{{ post.url }}">{{ post.title }}</a></h2>
      <p>{{ post.excerpt | strip_html }}</p>
    </li>
  {% endfor %}
</ul>
