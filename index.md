<h2>Posts:</h2>
<ul>
  {% for post in site.posts %}
    <li>
      <small>{{ post.date | date: '%B %d, %Y' }}</small>
      <h2><a href="{{ post.url }}">{{ post.title }}</a></h2>
      <p>{{ post.content }}</p>
    </li>
  {% endfor %}
</ul>
